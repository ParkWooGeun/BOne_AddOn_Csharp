using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작업일보등록(공정)
    /// </summary>
    internal class PS_PP041 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.DBDataSource oDS_PS_PP041H; // 등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP041L; // 등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP041M; // 등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP041N; // 등록라인
        private string oLastItemUID01; // 클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; // 마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; // 마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oMat01Row01;
        private int oMat02Row02;
        private int oMat03Row03;
        private string oDocType01;
        private string oDocEntry01;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP041.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP041_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP041");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                PS_PP041_CreateItems();
                PS_PP041_ComboBox_Setting();
                PS_PP041_EnableMenus();
                PS_PP041_SetDocument(oFormDocEntry);
                PS_PP041_FormResize();
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
        private void PS_PP041_CreateItems()
        {
            try
            {
                oDS_PS_PP041H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
                oDS_PS_PP041L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
                oDS_PS_PP041M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
                oDS_PS_PP041N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

                oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");

                oForm.DataSources.UserDataSources.Add("SOrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SOrdGbn").Specific.DataBind.SetBound(true, "", "SOrdGbn");

                oForm.DataSources.UserDataSources.Add("SCpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("SCpCode").Specific.DataBind.SetBound(true, "", "SCpCode");

                oForm.DataSources.UserDataSources.Add("SCpName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("SCpName").Specific.DataBind.SetBound(true, "", "SCpName");

                oForm.DataSources.UserDataSources.Add("SMoldNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("SMoldNo").Specific.DataBind.SetBound(true, "", "SMoldNo");

                oForm.DataSources.UserDataSources.Add("SMoldNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("SMoldNm").Specific.DataBind.SetBound(true, "", "SMoldNm");

                oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");

                oForm.DataSources.UserDataSources.Add("ilboChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ilboChk").Specific.DataBind.SetBound(true, "", "ilboChk");

                oDocType01 = "작업일보등록(공정)";
                if (oDocType01 == "작업일보등록(작지)")
                {
                    oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oDocType01 == "작업일보등록(공정)")
                {
                    oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.Items.Item("Focus").Visible = false; //찾기 등으로 문서를 처음 열었을 경우 focus를 지정해줄 용도의 dummy 컨트롤
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP041_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "Mat02", "LTime", "Y", "Y");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "Mat02", "LTime", "N", "N");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("LTime"), "PS_PP041", "Mat02", "LTime", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "10", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "20", "PSMT지원");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "30", "외주");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "40", "실적");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "50", "일반조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "OrdType", "", "60", "외주조정");
                dataHelpClass.Combo_ValidValues_SetValueItem((oForm.Items.Item("OrdType").Specific), "PS_PP041", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP041", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem((oForm.Items.Item("DocType").Specific), "PS_PP041", "DocType", false);

                oForm.Items.Item("SOrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SOrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN('104','107') order by Code", "", false, false);

                oForm.Items.Item("SBPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN('104','107') order by Code", "", false, false);

                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP041_EnableMenus()
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
        private void PS_PP041_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP041_FormItemEnabled();
                    PS_PP041_AddMatrixRow01(0, true);
                    PS_PP041_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP041_FormItemEnabled();
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
        private void PS_PP041_FormItemEnabled()
        {
            string query01;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                    oForm.Items.Item("UseMCode").Enabled = true;
                    oForm.Items.Item("MoldCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oMat02.Columns.Item("NStart").Editable = false; //비가동시작시간
                    oMat02.Columns.Item("NEnd").Editable = false; //비가동종료시간
                    oMat02.Columns.Item("NTime").Editable = true; //비가동시간만
                    oForm.Items.Item("Mat03").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;

                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //항상일반타입이여야함
                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //선택인 상태
                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("SOrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("SBPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.DataSources.UserDataSources.Item("SCpCode").Value = "";
                    oForm.Items.Item("UseMCode").Specific.Value = "";
                    oForm.DataSources.UserDataSources.Item("SMoldNo").Value = "";
                    oForm.Items.Item("MoldCode").Specific.Value = "";

                    PS_PP041_FormClear();
                    
                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else if (oDocType01 == "작업일보등록(공정)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
                    oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("UseMCode").Enabled = true;
                    oForm.Items.Item("MoldCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Mat03").Enabled = false;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oDS_PS_PP041H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
                    {
                        oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("UseMCode").Enabled = false;
                        oForm.Items.Item("MoldCode").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("Mat02").Enabled = false;
                        oForm.Items.Item("Mat03").Enabled = false;
                        oForm.Items.Item("Button01").Enabled = false;
                        oForm.Items.Item("1").Enabled = false;
                    }
                    else
                    {
                        if (oDS_PS_PP041H.GetValue("U_OrdType", 0).ToString().Trim() == "20") //PSMT
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("UseMCode").Enabled = false;
                            oForm.Items.Item("MoldCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP041H.GetValue("U_OrdType", 0).ToString().Trim() == "30") //외주
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("UseMCode").Enabled = false;
                            oForm.Items.Item("MoldCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP041H.GetValue("U_OrdType", 0).ToString().Trim() == "40") //실적
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("UseMCode").Enabled = false;
                            oForm.Items.Item("MoldCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP041H.GetValue("U_OrdType", 0).ToString().Trim() == "10" || oDS_PS_PP041H.GetValue("U_OrdType", 0).ToString().Trim() == "50") //일반,조정
                        {
                            if (oDS_PS_PP041H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104") //멀티
                            {
                                //MG 작업일보일 경우
                                // 1. 작업일보 실적 관리 여부(필요 : X)
                                // 2. 기준작업지시등록 문서 생산완료 등록 여부(필요 : O)
                                // 3. 다음 공정의 작업일보 등록 여부(필요 : O)
                                // 4. 수정 가능 권한 보유 여부(필요 : O)
                                // 5. V-MILL 공정(필요 : O)
                                // 6. 위 검사 결과 따라 화면 컨트롤 Enable 설정
                                // 7. 위 내용들을 반복문을 통해서 구현 -> 행의 수만큼 조회시간 증가
                                // 8. 로직 변경 : 위 5가지 조건 검사를 한번에 하거나 필요 없는 조건은 pass 하도록 수정

                                // ※ MG일경우 "다음 공정의 작업일보가 등록된 경우", "마지막 공정일 때 생산완료가 등록된 경우", "수정 가능 권한 여부", "V-MILL 공정 여부"에 따라 화면 컨트롤 Enable 설정 함

                                string superUserYN = dataHelpClass.GetValue("SELECT U_UseYN FROM [@PS_SY001L] A WHERE A.Code = 'A007' AND A.U_Minor = 'PS_PP041' AND A.U_RelCd = '" + PSH_Globals.oCompany.UserName + "'", 0, 1); //수정 가능 권한 보유 여부

                                query01 = "EXEC PS_PP041_50 '" + oDS_PS_PP041H.GetValue("DocEntry", 0) + "'"; //수정가능여부 프로시저
                                RecordSet01.DoQuery(query01);

                                if 
                                (
                                    oDS_PS_PP041H.GetValue("U_CpCode", 0).ToString().Trim() == "CP50101" //V-MILL 공정
                                    || RecordSet01.Fields.Item("RESULT").Value == "FALSE" //수정가능여부
                                ) //수정불가(컨트롤 Enable = false) 조건 체크
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("DocEntry").Enabled = false;
                                    oForm.Items.Item("UseMCode").Enabled = false;
                                    oForm.Items.Item("MoldCode").Enabled = false;
                                    oForm.Items.Item("DocDate").Enabled = false;
                                    oForm.Items.Item("Mat01").Enabled = false;
                                    oForm.Items.Item("Mat02").Enabled = false;
                                    oForm.Items.Item("Mat03").Enabled = false;
                                    oForm.Items.Item("Button01").Enabled = false;

                                    if (string.IsNullOrEmpty(superUserYN))
                                    {
                                        oForm.Items.Item("1").Enabled = false;
                                    }
                                    else if (superUserYN == "Y")
                                    {
                                        oForm.Items.Item("1").Enabled = true;
                                    }
                                }
                                else
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("DocEntry").Enabled = false;
                                    oForm.Items.Item("UseMCode").Enabled = true;
                                    oForm.Items.Item("MoldCode").Enabled = false;
                                    oForm.Items.Item("DocDate").Enabled = true;
                                    oForm.Items.Item("Mat01").Enabled = true;
                                    oForm.Items.Item("Mat02").Enabled = true;
                                    oForm.Items.Item("Mat03").Enabled = true;
                                    oForm.Items.Item("Button01").Enabled = true;
                                    oForm.Items.Item("1").Enabled = true;
                                }
                            }

                            oMat01.Columns.Item("BQty").Visible = true;
                            oMat01.Columns.Item("PSum").Visible = false;
                            oMat01.Columns.Item("PWeight").Visible = false;
                            oMat01.Columns.Item("YWeight").Visible = false;
                            oMat01.Columns.Item("NWeight").Visible = false;
                        }
                        else //멀티를 제외한 경우, 엔드베어링
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("UseMCode").Enabled = true;
                            oForm.Items.Item("MoldCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = true;
                            oForm.Items.Item("Mat02").Enabled = true;
                            oForm.Items.Item("Mat03").Enabled = true;
                            oForm.Items.Item("Button01").Enabled = true;
                            oForm.Items.Item("1").Enabled = true;

                            oMat01.Columns.Item("BQty").Visible = false;
                            oMat01.Columns.Item("PSum").Visible = true;
                            oMat01.Columns.Item("PWeight").Visible = true;
                            oMat01.Columns.Item("YWeight").Visible = true;
                            oMat01.Columns.Item("NWeight").Visible = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP041_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "");
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
        /// 메트릭스 Row추가(Mat01)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_PP041_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP041L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP041L.Offset = oRow;
                oDS_PS_PP041L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP041_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP041M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP041M.Offset = oRow;
                oDS_PS_PP041M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP041_AddMatrixRow03(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP041N.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP041N.Offset = oRow;
                oDS_PS_PP041N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP041_FormResize()
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
                oForm.Items.Item("Button02").Left = 365;
                oForm.Items.Item("Button02").Top = oForm.Items.Item("Opt03").Top;

                oForm.Items.Item("ilboChk").Left = 100;
                oForm.Items.Item("ilboChk").Top = oForm.Items.Item("Opt03").Top;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
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
        private bool PS_PP041_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            int j;
            double FailQty;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP041_FormClear();
                }

                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (Convert.ToDouble(oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "작업시간은 필수입니다.";
                        oMat02.Columns.Item("YTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                }

                if (dataHelpClass.Future_Date_Check(oForm.Items.Item("DocDate").Specific.Value) == "N")
                {
                    errMessage = "미래일자는 입력할 수 없습니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10" & oForm.Items.Item("OrdType").Specific.Selected.Value != "50")
                {
                    errMessage = "작업타입이 일반,조정이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                {
                    errMessage = "작업구분이 선택되지 않았습니다.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "공정정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                if (oMat02.VisualRowCount == 1)
                {
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "107")
                    {
                        oMat02.FlushToDataSource();
                        oDS_PS_PP041M.SetValue("U_WorkCode", 0, "9999999");
                        oDS_PS_PP041M.SetValue("U_WorkName", 0, "조정");
                        oDS_PS_PP041M.SetValue("U_YTime", 0, "1");
                        PS_PP041_AddMatrixRow02(1, false);
                        oMat02.LoadFromDataSource();
                    }
                    else
                    {
                        errMessage = "작업자정보 라인이 존재하지 않습니다.";
                        oMat02.SelectRow(oMat02.VisualRowCount, true, false);
                        throw new Exception();
                    }
                }

                if (oMat03.VisualRowCount == 0)
                {
                    errMessage = "불량정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                //마감상태 체크
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다." + (char)13 + "작업일보일자를 확인하고, 회계부서로 문의하세요.";
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
                    if (Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "생산수량또는 불량수량 또는 스크랩이 없습니다.";
                        oMat01.Columns.Item("YQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    //멀티 설비코드는 필수 입력
                    if (oForm.Items.Item("SOrdGbn").Specific.Value.ToString().Trim() == "104")
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("MachCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "설비코드는 필수입니다.";
                            oMat01.Columns.Item("MachCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    //MG생산 생산수량, 불량수량 없이 스크랩만 발생할 수 없습니다.
                    if (oForm.Items.Item("SOrdGbn").Specific.Value.ToString().Trim() == "104")
                    {
                        if 
                        (
                            Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value) == 0 
                         && Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) == 0 
                         && Convert.ToDouble(oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value) > 0
                        )
                        {
                            errMessage = "생산수량, 불량수량없이 스크랩이 발생할 수 없습니다.";
                            oMat01.Columns.Item("YQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    //엔드베어링은 생산수량 = 합격수량 + 불량수량 이어야 한다.
                    if (oForm.Items.Item("SOrdGbn").Specific.Value.ToString().Trim() == "107")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) != Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "생산수량 = 합격수량 + 불량수량 이어야 합니다.";
                            oMat01.Columns.Item("YQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)
                    {
                        //멀티 포장공정외는 공수필수
                        if (oForm.Items.Item("SOrdGbn").Specific.Value.ToString().Trim() == "104" && oForm.Items.Item("SCpCode").Specific.Value.ToString().Trim() != "CP50107")
                        {
                            errMessage = "실동시간은 필수입니다.";
                            oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    FailQty = 0;
                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        //불량코드를 입력했는지 check
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 & string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
                        {
                            errMessage = "불량수량이 입력되었을 때는 불량코드는 필수입니다.";
                            oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }

                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value)
                        {
                            FailQty += Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value);
                        }
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) != FailQty)
                    {
                        errMessage = "공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.";
                        throw new Exception();
                    }
                }

                //비가동코드와 비가동시간 체크
                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동코드가 입력되었을 때 비가동시간은 필수입니다.";
                            oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동시간이 입력되었을 때 비가동코드는 필수입니다.";
                            oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                }

                //불량정보 입력 체크
                if (string.IsNullOrEmpty(dataHelpClass.GetValue("select U_UseYN from [@PS_SY001L] a where a.Code ='A007' and a.U_Minor ='PS_PP041' and a.U_RelCd = '" + PSH_Globals.oCompany.UserName + "'", 0, 1))) //슈퍼유저가 아니면 불량정보 필수 입력
                {
                    for (i = 1; i <= oMat03.VisualRowCount - 1; i++)
                    {
                        //해당 작업지시의 재작업 여부 조회
                        if (dataHelpClass.GetValue("SELECT U_ReWorkYN FROM [@PS_PP030M] WHERE Convert(Nvarchar(50),DocEntry) +" + "'-'" + "+ Convert(Nvarchar(20),U_LineId) = '" + oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                        {
                            //불량코드
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(i).Specific.Value))
                            {
                                errMessage = "재작업 시 불량정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }

                            //원인공정코드
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("CsCpCode").Cells.Item(i).Specific.Value))
                            {
                                errMessage = "재작업 시 원인공정정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("CsCpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }

                            //작업자코드
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("CsWkCode").Cells.Item(i).Specific.Value))
                            {
                                errMessage = "재작업 시 작업자정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("CsWkCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }
                }

                if (PS_PP041_Validate("검사01") == false)
                {
                    errMessage = " ";
                    throw new Exception();
                }

                oDS_PS_PP041L.RemoveRecord(oDS_PS_PP041L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP041M.RemoveRecord(oDS_PS_PP041M.Size - 1);
                oMat02.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage == " ")
                {
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
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP041_Validate(string ValidateType)
        {
            bool returnValue = false;
            int i;
            int j;
            string query01;
            double prevDBCpQty;
            double prevMATRIXCpQty;
            double currentDBCpQty;
            double currentMATRIXCpQty;
            string prevCpInfo;
            string currentCpInfo;
            string nextCpInfo;
            DateTime prevDate;
            string ordMgNum;
            bool exist;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {   
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                    throw new Exception();
                }

                string ordType = oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim();

                if (ordType == "10" || ordType == "50") //작업타입이 일반,조정인경우
                {
                }
                else if (ordType == "20") //작업타입이 PSMT지원인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();
                }
                else if (ordType == "30") //작업타입이 외주인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();                    
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();
                }

                if (ValidateType == "검사01")
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //입력된 행에 대해
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                        {
                            errMessage = "작업지시문서가 존재하지 않습니다.";
                            throw new Exception();
                        }
                    }

                    if (ordType == "10") //작업타입이 일반인경우
                    {
                        //삭제된 행에 대한처리
                        query01 = "  SELECT     PS_PP040H.DocEntry,";
                        query01 += "            PS_PP040L.LineId,";
                        query01 += "            CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                        query01 += "            PS_PP040L.U_OrdGbn AS OrdGbn,";
                        query01 += "            PS_PP040L.U_PP030HNo AS PP030HNo,";
                        query01 += "            PS_PP040L.U_PP030MNo AS PP030MNo, ";
                        query01 += "            PS_PP040L.U_ordMgNum AS ordMgNum ";
                        query01 += " FROM       [@PS_PP040H] PS_PP040H";
                        query01 += "            LEFT JOIN";
                        query01 += "            [@PS_PP040L] PS_PP040L";
                        query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                        query01 += " WHERE      PS_PP040H.Canceled = 'N'";
                        query01 += "            AND PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                        RecordSet01.DoQuery(query01);

                        for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                        {
                            exist = false;
                            //기존에 있는 행에대한처리
                            for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                {
                                    //새로추가된 행인경우 검사 불필요
                                }
                                else
                                {
                                    if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value)
                                     && Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value)) //라인번호가 같고 문서번호가 같으면 존재하는행
                                    {
                                        exist = true;
                                    }
                                }
                            }

                            //삭제된 행중 수량관계를 알아본다.
                            if (exist == false)
                            {
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") //생산완료로 등록하는 실적 여부인지 조회
                                {
                                    //실적, 문서의 타입 필요
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "삭제된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }

                                if (RecordSet01.Fields.Item("OrdGbn").Value == "104") //멀티
                                {
                                    //다음공정이 존재하면
                                    nextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1);
                                    if (!string.IsNullOrEmpty(nextCpInfo))
                                    {
                                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + nextCpInfo + "'", 0, 1)) > 0)
                                        {
                                            errMessage = "후행공정이 입력된 문서입니다.";
                                            throw new Exception();
                                        }
                                    }
                                    else
                                    {
                                        //다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
                                    }
                                }
                                else if (RecordSet01.Fields.Item("OrdGbn").Value == "107") //엔드베어링
                                {
                                    ordMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
                                    currentCpInfo = ordMgNum;

                                    prevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + ordMgNum + "'", 0, 1);
                                    if (string.IsNullOrEmpty(prevCpInfo))
                                    {
                                        //해당공정이 첫공정이면 입력 무관
                                    }
                                    else
                                    {
                                        prevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + prevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                        prevMATRIXCpQty = 0;
                                        for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                        {
                                            if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == prevCpInfo)
                                            {
                                                prevMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                            }
                                        }

                                        currentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + currentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                        currentMATRIXCpQty = 0;
                                        for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                        {
                                            if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == currentCpInfo)
                                            {
                                                currentMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                            }
                                        }
                                        if ((prevDBCpQty + prevMATRIXCpQty) < (currentDBCpQty + currentMATRIXCpQty))
                                        {
                                            errMessage = "삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                            RecordSet01.MoveNext();
                        }

                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value))
                            {
                                //새로추가된 행인경우, 검사 불필요
                            }
                            else
                            {
                                if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "104") //멀티
                                {
                                    //수량이 수정되면 뒷공정이 존재한다면 수정할 수 없다.
                                    nextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1);
                                    if (!string.IsNullOrEmpty(nextCpInfo))
                                    {
                                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + nextCpInfo + "'", 0, 1)) > 0)
                                        {
                                            //다음공정으로 생성된 실적조회 만약 존재한다면 취소불가능
                                            //작업일보등록된문서중에 수정이 된문서를 구함
                                            query01 = "  SELECT     PS_PP040L.U_ordMgNum,";
                                            query01 += "            PS_PP040L.U_Sequence,";
                                            query01 += "            PS_PP040L.U_CpCode,";
                                            query01 += "            PS_PP040L.U_ItemCode,";
                                            query01 += "            PS_PP040L.U_PP030HNo,";
                                            query01 += "            PS_PP040L.U_PP030MNo,";
                                            query01 += "            PS_PP040L.U_PQty,";
                                            query01 += "            PS_PP040L.U_NQty,";
                                            query01 += "            PS_PP040L.U_ScrapWt";
                                            query01 += " FROM       [@PS_PP040H] PS_PP040H";
                                            query01 += "            LEFT JOIN";
                                            query01 += "            [@PS_PP040L] PS_PP040L";
                                            query01 += "                 ON PS_PP040H.DocEntry = PS_PP040L.DocEntry";
                                            query01 += " WHERE      PS_PP040H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                            query01 += "            AND PS_PP040L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                            query01 += "            AND PS_PP040H.Canceled = 'N'";
                                            RecordSet01.DoQuery(query01);

                                            if (RecordSet01.Fields.Item(0).Value.ToString().Trim() == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(1).Value.ToString().Trim() == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(2).Value.ToString().Trim() == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(3).Value.ToString().Trim() == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(4).Value.ToString().Trim() == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(5).Value.ToString().Trim() == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value
                                             && Convert.ToDouble(RecordSet01.Fields.Item(6).Value.ToString().Trim()) == Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value)
                                             && Convert.ToDouble(RecordSet01.Fields.Item(7).Value.ToString().Trim()) == Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value)
                                             && Convert.ToDouble(RecordSet01.Fields.Item(8).Value.ToString().Trim()) == Convert.ToDouble(oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value))
                                            {
                                                //값이 변경된 행의경우
                                            }
                                            else
                                            {
                                                errMessage = "후행공정이 입력된 문서입니다. 수정할 수 없습니다.";
                                                throw new Exception();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
                                    }

                                    //실적포인트면
                                    //현재 공정이 바렐 앞공정이면
                                    if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oForm.Items.Item("DocEntry").Specific.Value + "-" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                                    {
                                        //실적테이블에 값이 존재하는지 검사
                                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                                        {
                                            //작업일보등록된문서중에 수정이 된문서를 구함
                                            query01 = "  SELECT     PS_PP040L.U_ordMgNum,";
                                            query01 += "            PS_PP040L.U_Sequence,";
                                            query01 += "            PS_PP040L.U_CpCode,";
                                            query01 += "            PS_PP040L.U_ItemCode,";
                                            query01 += "            PS_PP040L.U_PP030HNo,";
                                            query01 += "            PS_PP040L.U_PP030MNo,";
                                            query01 += "            PS_PP040L.U_PQty,";
                                            query01 += "            PS_PP040L.U_NQty,";
                                            query01 += "            PS_PP040L.U_ScrapWt,";
                                            query01 += "            PS_PP040L.U_WorkTime";
                                            query01 += "            FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry";
                                            query01 += " WHERE      PS_PP040H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                            query01 += "            AND PS_PP040L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                            query01 += "            AND PS_PP040H.Canceled = 'N'";
                                            RecordSet01.DoQuery(query01);

                                            if (RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(3).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(4).Value == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(5).Value == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(6).Value == oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(7).Value == oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(8).Value == oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(9).Value == oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value)
                                            {
                                                //값이 변경된 행의경우
                                            }
                                            else
                                            {
                                                errMessage = "생산실적이 등록된 행은 수정할 수 없습니다.";
                                                throw new Exception();
                                            }
                                        }
                                    }
                                }
                                else if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "107") //엔드베어링
                                {
                                    //실적포인트면
                                    //현재 공정이 바렐 앞공정
                                    if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oForm.Items.Item("DocEntry").Specific.Value + "-" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                                    {
                                        //휘팅벌크포장,휘팅실적
                                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0
                                        || (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0))
                                        {
                                            //작업일보등록된문서중에 수정이 된문서를 구함
                                            query01 = "  SELECT     PS_PP040L.U_ordMgNum,";
                                            query01 += "            PS_PP040L.U_Sequence,";
                                            query01 += "            PS_PP040L.U_CpCode,";
                                            query01 += "            PS_PP040L.U_ItemCode,";
                                            query01 += "            PS_PP040L.U_PP030HNo,";
                                            query01 += "            PS_PP040L.U_PP030MNo,";
                                            query01 += "            PS_PP040L.U_PQty,";
                                            query01 += "            PS_PP040L.U_NQty,";
                                            query01 += "            PS_PP040L.U_ScrapWt,";
                                            query01 += "            PS_PP040L.U_WorkTime";
                                            query01 += " FROM       [@PS_PP040H] PS_PP040H";
                                            query01 += "            LEFT JOIN";
                                            query01 += "            [@PS_PP040L] PS_PP040L";
                                            query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry";
                                            query01 += " WHERE      PS_PP040H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                            query01 += "            AND PS_PP040L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                            query01 += "            AND PS_PP040H.Canceled = 'N'";
                                            RecordSet01.DoQuery(query01);

                                            if (RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(3).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(4).Value == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(5).Value == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(6).Value == oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(7).Value == oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(8).Value == oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value
                                             && RecordSet01.Fields.Item(9).Value == oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value)
                                            {
                                                //값이 변경된 행의경우
                                            }
                                            else
                                            {
                                                errMessage = "생산실적이 등록된 행은 수정할 수 없습니다.";
                                                throw new Exception();
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //입력된 모든행에 대해 입력가능성 검사
                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "104") //멀티
                            {
                                //상단에서 후행공정입력여부검사 및 실적등록여부 검사 수정되었다면 앞공정보다 많게는 입력할수 없다(Validate에서처리)
                            }
                            else if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "107") //엔드베어링
                            {
                                ordMgNum = oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value;
                                currentCpInfo = oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value;
                                prevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + ordMgNum + "'", 0, 1);

                                if (string.IsNullOrEmpty(prevCpInfo))
                                {
                                    //해당공정이 첫공정이면 입력되어도 상관없다.
                                }
                                else
                                {
                                    prevDate = dataHelpClass.GetValue("SELECT Max(PS_PP040H.U_DocDate) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + prevCpInfo + "' AND PS_PP040H.DocEntry <> '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1);

                                    if (oForm.Items.Item("DocDate").Specific.Value < prevDate)
                                    {
                                        errMessage = "현공정의 일자가 선행공정의 일자보다 빠릅니다. 확인바랍니다.";
                                        oMat01.SelectRow(i, true, false);
                                        throw new Exception();
                                    }

                                    prevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + prevCpInfo + "' AND PS_PP040H.DocEntry <> '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    prevMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == prevCpInfo)
                                        {
                                            prevMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }

                                    currentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + currentCpInfo + "' AND PS_PP040L.DocEntry <> '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    currentMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == currentCpInfo)
                                        {
                                            currentMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }

                                    if ((prevDBCpQty + prevMATRIXCpQty) < (currentDBCpQty + currentMATRIXCpQty))
                                    {
                                        errMessage = "선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.";
                                        oMat01.SelectRow(i, true, false);
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT인경우
                    {
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                    {
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                    {
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경우
                    {
                    }
                }
                else if (ValidateType == "행삭제01") //행삭제전 행삭제가능여부검사
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {   
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))
                        {
                            //새로추가된 행인경우 삭제가능
                        }
                        else
                        {   
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "107") //엔드베어링
                            {
                                //실적포인트이면
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    //휘팅벌크포장
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            else if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "104") //멀티
                            {
                                //실적포인트이면
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    query01 = "   SELECT    COUNT(*) ";
                                    query01 += "  FROM      [@PS_PP080H] PS_PP080H";
                                    query01 += "            LEFT JOIN";
                                    query01 += "            [@PS_PP080L] PS_PP080L";
                                    query01 += "                ON PS_PP080H.DocEntry = PS_PP080L.DocEntry";
                                    query01 += "  WHERE     ISNULL(PS_PP080L.U_OIGENum,'') = ''";
                                    query01 += "            AND PS_PP080L.U_PP030HNo = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "'";
                                    query01 += "            AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'";
                                    query01 += "            AND ISNULL(PS_PP080L.U_Check, 'N') = 'N'";

                                    if (Convert.ToInt32(dataHelpClass.GetValue(query01, 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }

                                //후행공정으로 등록한 작업일보가 있으면 수정 불가
                                nextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1);
                                if (!string.IsNullOrEmpty(nextCpInfo))
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + nextCpInfo + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "후행공정이 존재합니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경우
                    {
                    }
                }
                else if (ValidateType == "수정01") //수정전 수정가능여부검사
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {   
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))
                        {
                            //새로추가된 행인경우 수정가능
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "107") //엔드베어링
                            {
                                //실적포인트이면
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    //휘팅벌크포장
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            else if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "104") //멀티이면서
                            {
                                //실적포인트이면
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oForm.Items.Item("DocEntry").Specific.Value + "-" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    //실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }

                                //후행공정으로 등록한 작업일보가 있으면 수정 불가
                                nextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1);
                                if (!string.IsNullOrEmpty(nextCpInfo))
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + nextCpInfo + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "후행공정이 존재합니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT인경우
                    {
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                    {
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경우
                    {
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errMessage = "이미취소된 문서 입니다. 취소할 수 없습니다.";
                        throw new Exception();
                    }
                    
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
                        //삭제된 행에 대한처리
                        query01 = "  SELECT     PS_PP040H.DocEntry,PS_PP040L.LineId,";
                        query01 += "            CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                        query01 += "            PS_PP040L.U_OrdGbn AS OrdGbn,";
                        query01 += "            PS_PP040L.U_PP030HNo AS PP030HNo,";
                        query01 += "            PS_PP040L.U_PP030MNo AS PP030MNo, ";
                        query01 += "            PS_PP040L.U_ordMgNum AS ordMgNum ";
                        query01 += " FROM       [@PS_PP040H] PS_PP040H";
                        query01 += "            LEFT JOIN";
                        query01 += "            [@PS_PP040L] PS_PP040L";
                        query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                        query01 += " WHERE      PS_PP040H.Canceled = 'N'";
                        query01 += "            AND PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                        RecordSet01.DoQuery(query01);

                        for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                        {   
                            //마지막공정에 항상 실적포인트가 되어있음
                            if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y")
                            {
                                //실적, 문서의타입필요
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "' AND PS_PP080H.Status = 'O'", 0, 1)) > 0)
                                {
                                    errMessage = "생산실적 등록된 문서입니다. 적용할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                            
                            if (RecordSet01.Fields.Item("OrdGbn").Value == "104") //멀티
                            {
                                //다음공정이 존재하면
                                nextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1);
                                if (!string.IsNullOrEmpty(nextCpInfo))
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + nextCpInfo + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "후행공정이 입력된 문서입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                                else
                                {
                                    //다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
                                }
                            }
                            else if (RecordSet01.Fields.Item("OrdGbn").Value == "107") //엔드베어링
                            {
                                //삭제된 행에 대한 검사
                                ordMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
                                currentCpInfo = ordMgNum;

                                prevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + ordMgNum + "'", 0, 1);
                                if (string.IsNullOrEmpty(prevCpInfo))
                                {
                                    //해당공정이 첫공정이면 입력 가능
                                }
                                else
                                {
                                    prevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + prevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    prevMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == prevCpInfo)
                                        {
                                            prevMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }
                                    currentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_ordMgNum = '" + currentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    currentMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == currentCpInfo)
                                        {
                                            currentMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }
                                    if ((prevDBCpQty + prevMATRIXCpQty) < (currentDBCpQty + currentMATRIXCpQty))
                                    {
                                        errMessage = "취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            RecordSet01.MoveNext();
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                    {   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경우
                    {
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
            }

            return returnValue;
        }

        /// <summary>
        /// OrderInfoLoad
        /// </summary>
        private void PS_PP041_OrderInfoLoad()
        {
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                
                if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입 일반,조정
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("SCpCode").Specific.Value)) //공정코드
                    {
                        errMessage = "작업지시 공정을 입력하지 않습니다.";
                        throw new Exception();
                    }
                    else
                    {
                        //작업구분선택
                        oForm.Items.Item("CpCode").Specific.Value = oForm.Items.Item("SCpCode").Specific.Value.ToString().Trim();
                        oForm.Items.Item("CpName").Specific.Value = oForm.Items.Item("SCpName").Specific.Value.ToString().Trim();
                        
                        if (oForm.Items.Item("SOrdGbn").Specific.Selected.Value == "선택") //값이 선택되어 있지 않다면 기본으로
                        {
                            oForm.Items.Item("OrdGbn").Specific.Select(dataHelpClass.GetValue("SELECT U_ItmBsort FROM [@PS_PP001L] WHERE U_CpCode = '" + oForm.Items.Item("SCpCode").Specific.Value + "'", 0, 1), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else //값이 선택되었다면 선택된 값으로..
                        {
                            oForm.Items.Item("OrdGbn").Specific.Select(oForm.Items.Item("SOrdGbn").Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        
                        if (oForm.Items.Item("SBPLId").Specific.Selected.Value == "선택") //사업장
                        {
                            oForm.Items.Item("BPLId").Specific.Select(oForm.Items.Item("SBPLId").Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else
                        {
                            oForm.Items.Item("BPLId").Specific.Select(oForm.Items.Item("SBPLId").Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }

                        //공정코드에 따라 매트릭스 변경
                        if (dataHelpClass.GetValue("SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = '" + oForm.Items.Item("CpCode").Specific.Value + "'", 0, 1) == "CP101") //엔드베어링
                        {
                            oMat01.Columns.Item("BQty").Visible = false;
                            oMat01.Columns.Item("PSum").Visible = true;
                            oMat01.Columns.Item("PWeight").Visible = true;
                            oMat01.Columns.Item("YWeight").Visible = true;
                            oMat01.Columns.Item("NWeight").Visible = true;
                        }
                        else if (dataHelpClass.GetValue("SELECT Code FROM [@PS_PP001L] WHERE U_CpCode = '" + oForm.Items.Item("CpCode").Specific.Value + "'", 0, 1) == "CP501") //멀티
                        {
                            oMat01.Columns.Item("BQty").Visible = true;
                            oMat01.Columns.Item("PSum").Visible = false;
                            oMat01.Columns.Item("PWeight").Visible = false;
                            oMat01.Columns.Item("YWeight").Visible = false;
                            oMat01.Columns.Item("NWeight").Visible = false;
                        }

                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                        PS_PP041_AddMatrixRow01(0, true);

                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP041_AddMatrixRow02(0, true);

                        oMat03.Clear();
                        oMat03.FlushToDataSource();
                        oMat03.LoadFromDataSource();

                        oForm.Update();
                    }
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입 PSMT
                {
                    errMessage = "PSMT지원은 입력할 수 없습니다.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입 외주
                {
                    errMessage = "외주는 입력할 수 없습니다.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입 실적
                {
                    errMessage = "실적은 입력할 수 없습니다.";
                    throw new Exception();
                }
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FindValidateDocument
        /// </summary>
        /// <param name="ObjectType"></param>
        /// <returns></returns>
        private bool PS_PP041_FindValidateDocument(string ObjectType)
        {
            bool returnValue = false;
            string query01;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                
                query01 = "  SELECT DocEntry";
                query01 += " FROM   [" + ObjectType + "]";
                query01 += " WHERE  DocEntry = " + DocEntry;

                if (oDocType01 == "작업일보등록(작지)")
                {
                    query01 += " AND U_DocType = '10'";
                }
                else if (oDocType01 == "작업일보등록(공정)")
                {
                    query01 += " AND U_DocType = '20'";
                }
                RecordSet01.DoQuery(query01);

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
            catch(Exception ex)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 출고 DI(멀티,엔드베어링의 경우 첫공정이면서 처음 작업일보 등록시 투입자재를 출고 시킨다.)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP041_InsertoInventoryGenExit()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int RetVal;
            string afterDIDocNum;
            string sQry;
            int i;
            int oRow;
            int Cnt = 0;
            SAPbobsCOM.Documents DI_oInventoryGenExit = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

                for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                {
                    if (oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107") //멀티 또는 엔드베어링일경우
                    {
                        //107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 원자재가 스크랩을 이용이므로, 불출될 원자재가 없음.
                        if (oDS_PS_PP041L.GetValue("U_ItemCode", oRow).ToString().Trim() != "107010002" && oDS_PS_PP041L.GetValue("U_ItemCode", oRow).ToString().Trim() != "107010004")
                        {
                            if (oDS_PS_PP041L.GetValue("U_Sequence", oRow).ToString().Trim() == "1") //첫공정일 경우
                            {
                                sQry = "  select    b.docentry";
                                sQry += " from      [@PS_PP040L] a";
                                sQry += "           inner join";
                                sQry += "           [@PS_PP040H] b";
                                sQry += "               on a.docentry=b.docentry ";
                                sQry += " where     a.U_OrdGbn in ('104','107')";
                                sQry += "           and b.canceled <> 'Y' ";
                                sQry += "           and a.U_PP030HNo = '" + oDS_PS_PP041L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "' ";
                                sQry += "           and a.U_Sequence = '" + oDS_PS_PP041L.GetValue("U_Sequence", oRow).ToString().Trim() + "' ";
                                oRecordSet.DoQuery(sQry);

                                if (oRecordSet.RecordCount < 1) //처음 작업일보 등록시
                                {
                                    Cnt += 1;
                                }
                            }
                        }
                    }
                }

                if (Cnt < 1)
                {
                    returnValue = true;
                    errCode = "0";
                    throw new Exception();
                }
                
                DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); //문서타입(입고)

                i = 1;
                
                //Header
                DI_oInventoryGenExit.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                DI_oInventoryGenExit.TaxDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                DI_oInventoryGenExit.Comments = "작업일보등록(" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + ") 출고_PS_PP041";

                //Line
                for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                {
                    if (oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107") //멀티, 엔드베어링
                    {
                        //첫공정일 경우
                        if (oDS_PS_PP041L.GetValue("U_Sequence", oRow).ToString().Trim() == "1")
                        {
                            sQry = "  select    b.docentry";
                            sQry += " from      [@PS_PP040L] a";
                            sQry += "           inner join";
                            sQry += "           [@PS_PP040H] b";
                            sQry += "               on a.docentry = b.docentry";
                            sQry += " where     a.U_OrdGbn in ('104','107')";
                            sQry += "           and b.canceled <> 'Y' ";
                            sQry += "           and a.U_PP030HNo = '" + oDS_PS_PP041L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "' ";
                            sQry += "           and a.U_Sequence = '" + oDS_PS_PP041L.GetValue("U_Sequence", oRow).ToString().Trim() + "' ";
                            oRecordSet.DoQuery(sQry);

                            //처음 작업일보 등록시
                            if (oRecordSet.RecordCount < 1)
                            {
                                if (DI_oInventoryGenExit.Lines.Count < i)
                                {
                                    DI_oInventoryGenExit.Lines.Add();
                                    DI_oInventoryGenExit.Lines.BatchNumbers.Add();
                                }

                                sQry = "select U_ItemCode,U_ItemName,U_BatchNum,U_Weight from [@PS_PP030L] where docentry = '" + oDS_PS_PP041L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);

                                DI_oInventoryGenExit.Lines.SetCurrentLine(i - 1);
                                DI_oInventoryGenExit.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
                                DI_oInventoryGenExit.Lines.ItemDescription = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
                                DI_oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim();
                                DI_oInventoryGenExit.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
                                DI_oInventoryGenExit.Lines.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());

                                sQry = "  SELECT    TOP 1 T1.WhsCode ";
                                sQry += " FROM      [OBTN] AS T0";
                                sQry += "           LEFT JOIN";
                                sQry += "           [OBTQ] AS T1";
                                sQry += "               ON T0.ItemCode = T1.ItemCode";
                                sQry += "               AND T0.SysNumber = T1.SysNumber";
                                sQry += " WHERE     T0.DistNumber = '" + oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim() + "'";
                                sQry += "           AND T0.ItemCode = '" + oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim() + "'";

                                oRecordSet.DoQuery(sQry);
                                DI_oInventoryGenExit.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim();

                                i += 1;
                            }
                        }
                    }
                }

                RetVal = DI_oInventoryGenExit.Add();
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);
                    i = 1;
                    for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                    {
                        //멀티, 엔드베어링일경우
                        if (oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP041L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107")
                        {
                            //첫공정일 경우
                            if (oDS_PS_PP041L.GetValue("U_Sequence", oRow).ToString().Trim() == "1")
                            {
                                oDS_PS_PP041L.SetValue("U_OutDoc", oRow, afterDIDocNum);
                                oDS_PS_PP041L.SetValue("U_OutLin", oRow, Convert.ToString(i));

                                i += 1;
                            }
                        }
                    }
                }

                oMat01.LoadFromDataSource();
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
                else if (errCode == "0")
                {
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
        /// 입고DI(출고 취소)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP041_InsertoInventoryGenEntry()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int RetVal;
            string afterDIDocNum;
            string sQry;
            int i;
            int oRow;
            int Cnt = 0;
            SAPbobsCOM.Documents DI_oInventoryGenEntry = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

                for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                {
                    if (!string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_OutDoc", oRow).ToString().Trim())) //출고되어진 문서가 있는경우
                    {
                        Cnt += 1;
                    }
                }

                if (Cnt < 1)
                {
                    returnValue = true;
                    errCode = "0";
                    throw new Exception();
                }

                DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry); //문서타입(입고)

                i = 1;

                //Header
                DI_oInventoryGenEntry.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                DI_oInventoryGenEntry.TaxDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                DI_oInventoryGenEntry.Comments = "작업일보등록(" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + ") 출고 취소_PS_PP041";
                DI_oInventoryGenEntry.UserFields.Fields.Item("U_CtrlType").Value = "C";
                DI_oInventoryGenEntry.UserFields.Fields.Item("U_CancDoc").Value = oDS_PS_PP041L.GetValue("U_OutDoc", 0); //불출취소시 관리유형(취소) 원재료 입고현황과 구분을 하기위함

                //Line
                for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                {
                    if (!string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_OutDoc", oRow).ToString().Trim())) //출고 문서가 있는경우
                    {
                        if (DI_oInventoryGenEntry.Lines.Count < i)
                        {
                            DI_oInventoryGenEntry.Lines.Add();
                            DI_oInventoryGenEntry.Lines.BatchNumbers.Add();
                        }

                        sQry = "select U_ItemCode,U_ItemName,U_BatchNum,U_Weight from [@PS_PP030L] where docentry = '" + oDS_PS_PP041L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);

                        DI_oInventoryGenEntry.Lines.SetCurrentLine(i - 1);
                        DI_oInventoryGenEntry.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
                        DI_oInventoryGenEntry.Lines.ItemDescription = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
                        DI_oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim();
                        DI_oInventoryGenEntry.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
                        DI_oInventoryGenEntry.Lines.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());

                        //출고 창고 select
                        sQry = "  select    WhsCode";
                        sQry += " from      [IGE1]";
                        sQry += " where     docentry = '" + oDS_PS_PP041L.GetValue("U_OutDoc", oRow).ToString().Trim() + "' ";
                        sQry += "           and linenum = '" + (Convert.ToInt32(oDS_PS_PP041L.GetValue("U_OutLin", oRow).ToString().Trim()) - 1) + "'";
                        oRecordSet.DoQuery(sQry);
                        DI_oInventoryGenEntry.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim();

                        i += 1;
                    }
                }

                RetVal = DI_oInventoryGenEntry.Add();
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);
                    i = 1;
                    for (oRow = 0; oRow <= oDS_PS_PP041L.Size - 1; oRow++)
                    {
                        if (!string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_OutDoc", oRow).ToString().Trim())) //출고 문서가 있는경우
                        {
                            sQry = "  Update    [@PS_PP040L]";
                            sQry += " set       U_OutDocC = '" + afterDIDocNum + "',";
                            sQry += "           U_OutLinC = '" + i + "'";
                            sQry += " where     docentry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                            sQry += "           and visorder = '" + oRow + "'";
                            oRecordSet.DoQuery(sQry);

                            i += 1;
                        }
                    }
                }

                oMat01.LoadFromDataSource();
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
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
                else if (errCode == "0")
                {
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
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            double totTime = 0;
            double unitTime;
            double unitRemainTime;
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
                            if (PS_PP041_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //원재료 불출없이 추가시 주석 시작
                            if (PS_PP041_InsertoInventoryGenExit() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //원재료 불출없이 추가시 주석 종료

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP041_DataValidCheck() == false)
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
                    
                    if (pVal.ItemUID == "2") //취소버튼 누를시 저장할 자료가 있으면 메시지 표시
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oMat01.VisualRowCount > 1)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("저장하지 않은 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소") == 2)
                                {
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }

                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP041_OrderInfoLoad();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button02") //작업시간배부
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            for (int i = 1; i <= oMat02.VisualRowCount - 1; i++)
                            {
                                totTime += Convert.ToDouble(string.IsNullOrEmpty(oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value) == true ? "0" : oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value);
                            }

                            if (totTime > 0)
                            {
                                unitTime = Convert.ToDouble((totTime / (oMat01.VisualRowCount - 1)).ToString("#,##0.##"));
                                unitRemainTime = Convert.ToDouble((totTime - unitTime * (oMat01.VisualRowCount - 1)).ToString("#,##0.##"));

                                for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    if (i != oMat01.VisualRowCount - 2)
                                    {
                                        oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                    }
                                    else
                                    {
                                        oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                    }
                                }
                            }
                            oMat01.LoadFromDataSource();

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
                                PS_PP041_FormItemEnabled();
                                PS_PP041_AddMatrixRow02(0, true); //작업자 매트릭스 빈행 추가
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
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP041_FormItemEnabled();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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
            string errMessage = string.Empty;
            SAPbouiCOM.BoStatusBarMessageType messageType = BoStatusBarMessageType.smt_Error;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //일반,조정
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                                {
                                    errMessage = "작업구분이 선택되지 않았습니다.";
                                    messageType = BoStatusBarMessageType.smt_Warning;
                                    BubbleEvent = false;
                                    throw new Exception();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value))
                                {
                                    errMessage = "공정이 선택되지 않았습니다.";
                                    messageType = BoStatusBarMessageType.smt_Warning;
                                    BubbleEvent = false;
                                    throw new Exception();
                                }
                                else
                                {
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104")
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value == "")
                                        {
                                            PS_PP042 oTempClass = new PS_PP042();
                                            oTempClass.LoadForm(oForm, pVal.Row);
                                            BubbleEvent = false;
                                        }
                                    }
                                    else
                                    {
                                        dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
                                    }
                                }
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //지원
                            {
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //외주
                            {
                            } 
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //실적
                            {
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "WorkCode")
                        {
                            if (Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) == 0)
                            {
                                errMessage = "기준시간을 입력하지 않았습니다.";
                                messageType = BoStatusBarMessageType.smt_Warning;
                                oForm.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                BubbleEvent = false;
                                throw new Exception();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "SCpCode")
                    {
                        if (oForm.Items.Item("SOrdGbn").Specific.Selected.Value == "선택")
                        {
                            errMessage = "작업구분이 선택되지 않았습니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                    }
                    else if (pVal.ItemUID == "SMoldNo")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("SCpCode").Specific.Value))
                        {
                            errMessage = "공정이 선택되지 않았습니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oForm.Items.Item("SMoldNo").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "UseMCode")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("UseMCode").Specific.Value))
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "UseMCode", "");
                        }
                    }

                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "WorkCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "SCpCode", "");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "NCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "FailCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "CsCpCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "CsWkCode");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MoldNo");
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, messageType);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
                                oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP041_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP041M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP041_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                            }
                            else
                            {
                                oDS_PS_PP041N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "특정컬럼")
                            {
                                oDS_PS_PP041H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if (pVal.ItemUID == "SBPLId" || pVal.ItemUID == "SOrdGbn")
                            {
                            }
                            else
                            {
                                oDS_PS_PP041H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }

                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oMat02.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat03.LoadFromDataSource();
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat03";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
                    }
                    else if(pVal.ItemUID == "Opt03")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
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
                            if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 일반,조정
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                }
                                else
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP041_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP041_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }

                                    oDS_PS_PP041N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP041N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP041N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OrdMgNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OrdMgNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원인경우
                            {   
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                            {
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
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
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            PS_PP030 oTempClass = new PS_PP030();
                            oTempClass.LoadForm(codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")));
                        }
                        else if (pVal.ColUID == "PP030HNo")
                        {
                            PS_PP030 oTempClass = new PS_PP030();
                            oTempClass.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            PS_PP030 oTempClass = new PS_PP030();
                            oTempClass.LoadForm(codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")));
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
            double weight;
            double unitTime;
            double unitRemainTime;
            double time;
            double hour;
            double minute;
            string ordMgNum;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
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
                            if (PS_PP041_Validate("수정01") == false)
                            {
                                oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrdMgNum")
                                {
                                    ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택") //작업구분에 값이 없으면 작업지시가 불러오기전
                                    {
                                        oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                    }
                                    else //작업지시가 선택된상태
                                    {
                                        if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 일반,조정
                                        {
                                            if (oForm.Items.Item("BPLId").Specific.Value == "2")
                                            {
                                                //품질부적합 등록 여부 검사
                                                query01 = "EXEC PS_PP041_80 '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value + "'";
                                                RecordSet01.DoQuery(query01);
                                                if (RecordSet01.Fields.Item("Result").Value == "Y1")
                                                {
                                                    errMessage = "품질부적합 등록건(등록상태)입니다. 품질보증팀(검사자)에 문의하십시오.";
                                                    throw new Exception();
                                                }
                                                else if (RecordSet01.Fields.Item("Result").Value == "Y2")
                                                {
                                                    errMessage = "품질부적합 등록건(1차 해제상태)입니다. 품질보증팀(팀장)에 문의하십시오.";
                                                    throw new Exception();
                                                }
                                            }

                                            for (i = 1; i <= oMat01.RowCount; i++)
                                            {
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                                {
                                                    errMessage = "이미 입력한 공정입니다.";
                                                    oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                                    {
                                                        PS_PP041_AddMatrixRow01(pVal.Row, false);
                                                    }
                                                    throw new Exception();
                                                }
                                            }

                                            query01 = "EXEC PS_PP041_02 '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                            RecordSet01.DoQuery(query01);
                                            if (RecordSet01.RecordCount == 0)
                                            {
                                                oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else
                                            {
                                                oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP041L.SetValue("U_Sequence", pVal.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
                                                oDS_PS_PP041L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP041L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
                                                oDS_PS_PP041L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
                                                oDS_PS_PP041L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
                                                oDS_PS_PP041L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP041L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                                oDS_PS_PP041L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
                                                oDS_PS_PP041L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
                                                oDS_PS_PP041L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
                                                oDS_PS_PP041L.SetValue("U_BatchNum", pVal.Row - 1, RecordSet01.Fields.Item("BatchNum").Value);
                                                oDS_PS_PP041L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
                                                oDS_PS_PP041L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
                                                oDS_PS_PP041L.SetValue("U_PSum", pVal.Row - 1, RecordSet01.Fields.Item("PSum").Value);
                                                oDS_PS_PP041L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);

                                                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "104") //멀티
                                                {
                                                    if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50101") //vmill공정
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50102") //열처리공정
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50103") //pk공정
                                                    {
                                                        if (RecordSet01.Fields.Item("ReWorkYN").Value == "N")
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - 10));
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - 10));
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "10");
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                        }
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50104") //2차압연공정
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - 10));
                                                        oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - 10));
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "10");
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50105") //SLITTER공정
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50108") //P&F
                                                    {
                                                        if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "1")
                                                        {
                                                            //S&D에 생산 포장라벨 중량 자동 표시 (창원사업장)
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(Math.Round(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - Convert.ToDouble(RecordSet01.Fields.Item("PackWg").Value), 2))); //소수점 2째자리 반올림 처리
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                        }
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50109") //S&D
                                                    {
                                                        if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "1")
                                                        {
                                                            double neeee = RecordSet01.Fields.Item("BQty").Value;
                                                            //S&D에 생산 포장라벨 중량 자동 표시 (창원사업장)
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(Math.Round(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - Convert.ToDouble(RecordSet01.Fields.Item("PackWg").Value), 2))); //소수점 2째자리 반올림 처리
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                        }
                                                        //S&D(Slitter & DEGREASER) 공정(코드 : CP50109)에서 LOSS 발생시 최초투입중량의 0.6%, 추가(2017.01.02 송명규, 노근용 요청)
                                                        oDS_PS_PP041L.SetValue("U_LOSS", pVal.Row - 1, RecordSet01.Fields.Item("LOSS").Value);
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50106") //탈지공정
                                                    {
                                                        if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "1")
                                                        {
                                                            //탈지공정에 생산 포장라벨 중량 자동 표시 (창원사업장)
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(Math.Round(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - Convert.ToDouble(RecordSet01.Fields.Item("PackWg").Value), 2))); //소수점 2째자리 반올림 처리
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                        }
                                                        //탈지공정에서 LOSS발생 최초투입중량의 0.6%
                                                        oDS_PS_PP041L.SetValue("U_LOSS", pVal.Row - 1, RecordSet01.Fields.Item("LOSS").Value);
                                                    }
                                                    else if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50107") //PACKING공정
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value); //생산포장라벨중량 자동 표기(2016.07.18 송명규)
                                                        oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("PackWg").Value); //생산포장라벨중량 자동 표기(2016.07.18 송명규)
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(Math.Round(Convert.ToDouble(RecordSet01.Fields.Item("BQty").Value) - Convert.ToDouble(RecordSet01.Fields.Item("PackWg").Value), 2))); //소수점 2째자리 반올림 처리
                                                    }
                                                    else
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                        oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    }

                                                    oDS_PS_PP041L.SetValue("U_MachCode", pVal.Row - 1, RecordSet01.Fields.Item("MachCode").Value);
                                                    oDS_PS_PP041L.SetValue("U_MachName", pVal.Row - 1, RecordSet01.Fields.Item("MachName").Value);
                                                    oDS_PS_PP041L.SetValue("U_MoldNo", pVal.Row - 1, RecordSet01.Fields.Item("MoldNo").Value);
                                                    oDS_PS_PP041L.SetValue("U_MoldName", pVal.Row - 1, RecordSet01.Fields.Item("MoldName").Value);
                                                }
                                                else
                                                {
                                                    //엔드베어링 생산수량 구하기
                                                    ordMgNum = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
                                                    query01 = "EXEC [PS_PP041_03] '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                                    RecordSet01.DoQuery(query01);

                                                    oDS_PS_PP041L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item(0).Value);
                                                    oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item(0).Value);
                                                    oDS_PS_PP041L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                }

                                                oDS_PS_PP041L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP041L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP041L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP041L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP041L.SetValue("U_LineId", pVal.Row - 1, "");
                                                oDS_PS_PP041L.SetValue("U_WorkTime", pVal.Row - 1, "0");

                                                if (!string.IsNullOrEmpty(oForm.Items.Item("UseMCode").Specific.Value))
                                                {
                                                    oDS_PS_PP041L.SetValue("U_MachCode", pVal.Row - 1, oForm.Items.Item("UseMCode").Specific.Value);
                                                    oDS_PS_PP041L.SetValue("U_MachName", pVal.Row - 1, oForm.Items.Item("UseMName").Specific.Value);
                                                }
                                                else if (!string.IsNullOrEmpty(oForm.Items.Item("SMoldNo").Specific.Value)) //금형번호
                                                {
                                                    oDS_PS_PP041L.SetValue("U_MoldNo", pVal.Row - 1, oForm.Items.Item("SMoldNo").Specific.Value);
                                                    oDS_PS_PP041L.SetValue("U_MoldName", pVal.Row - 1, oForm.Items.Item("SMoldNm").Specific.Value);
                                                }

                                                //불량코드테이블
                                                if (oMat03.VisualRowCount == 0)
                                                {
                                                    PS_PP041_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    PS_PP041_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                }

                                                oDS_PS_PP041N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP041N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP041N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);

                                                if (oForm.Items.Item("SOrdGbn").Specific.Value.ToString().Trim() == "104")
                                                {
                                                    if (!string.IsNullOrEmpty(RecordSet01.Fields.Item("FailCode").Value.ToString().Trim()))
                                                    {
                                                        oDS_PS_PP041N.SetValue("U_FailCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("FailCode").Value);
                                                        oDS_PS_PP041N.SetValue("U_FailName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("FailName").Value);
                                                    }
                                                }
                                            }
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원
                                        {
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주
                                        {
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적
                                        {
                                        }

                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP041_AddMatrixRow01(pVal.Row, false);
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "PQty")
                                {
                                    if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "104") //멀티이면
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0"); //불량일경우
                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, "0"); //합격수량 0
                                            oDS_PS_PP041L.SetValue("U_NQty", pVal.Row - 1, oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value); //전량 불량처리
                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, "0"); //스크랩 0
                                        }
                                        else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value))
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                        else
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP041L.SetValue("U_NQty", pVal.Row - 1, "0"); //불량 0
                                            oDS_PS_PP041L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(Math.Round(Convert.ToDouble(oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value), 2)));
                                        }
                                    }
                                    else if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "107") //엔드베어링이면
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                        else
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                            if (weight == 0)
                                            {
                                                oDS_PS_PP041L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP041L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            }
                                            else
                                            {
                                                oDS_PS_PP041L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP041L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            oDS_PS_PP041L.SetValue("U_NQty", pVal.Row - 1, "0");
                                            oDS_PS_PP041L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "NQty")
                                {
                                    if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "104") //멀티이면
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                        }
                                        else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                        }
                                        else
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                    }
                                    else if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "107") //엔드베어링이면
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) < 0)
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                        else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP041L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                        else
                                        {
                                            oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP041L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                            if (weight == 0)
                                            {
                                                oDS_PS_PP041L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP041L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            else
                                            {
                                                oDS_PS_PP041L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP041L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                            }
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "WorkTime")
                                {
                                    oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                                else if (pVal.ColUID == "MachCode")
                                {
                                    oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP041L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "MoldNo")
                                {
                                    oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP041L.SetValue("U_MoldName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_Item + '[' + U_Callsize +']' FROM [@PS_PP190H] WHERE Code = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP041L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
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
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value); //기타작업
                                oDS_PS_PP041M.SetValue("U_WorkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP041M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP041_AddMatrixRow02(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "NStart")
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP041M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP041M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP041M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티일때만
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배, V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount == 1 ? 1 : (oMat01.VisualRowCount - 1));
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }

                                    hour = time / 100;
                                    minute = time % 100;
                                    time = hour;

                                    if (minute > 0)
                                    {
                                        time += 0.5;
                                    }

                                    oDS_PS_PP041M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP041M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP041M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티일때만
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배, V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount == 1 ? 1 : (oMat01.VisualRowCount - 1));
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (pVal.ColUID == "NEnd")
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP041M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP041M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP041M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티일때만
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배, V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount == 1 ? 1 : (oMat01.VisualRowCount - 1));
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }

                                    hour = time / 100;
                                    minute = time % 100;
                                    time = hour;

                                    if (minute > 0)
                                    {
                                        time += 0.5;
                                    }

                                    oDS_PS_PP041M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP041M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP041M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티일때만
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount == 1 ? 1 : (oMat01.VisualRowCount - 1));
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP041M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (pVal.ColUID == "YTime")
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP041M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));

                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티일때만
                                {
                                    if (oMat02.VisualRowCount > 1)
                                    {
                                        if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                        {
                                            //해당작지의 첫공정과 공정정보의 공정이 다르면 분배, V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                            if (Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0)
                                            {
                                                unitTime = Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / (oMat01.VisualRowCount == 1 ? 1 : (oMat01.VisualRowCount - 1));
                                                unitRemainTime = Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) - (unitTime * (oMat01.VisualRowCount - 1));
                                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                {
                                                    if (i != oMat01.VisualRowCount - 2)
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                    }
                                                    else
                                                    {
                                                        oDS_PS_PP041L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                oDS_PS_PP041M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }

                            oMat02.LoadFromDataSource();
                            oMat02.AutoResizeColumns();
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP041N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP041N.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ColUID == "CsCpCode")
                            {
                                oDS_PS_PP041N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP041N.SetValue("U_CsCpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ColUID == "CsWkCode") //원인공정 작업자 정보 추가(2012.03.20 송명규)
                            {
                                oDS_PS_PP041N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP041N.SetValue("U_CsWkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT T0.lastName+T0.firstName FROM OHEM T0 WHERE T0.U_MSTCOD = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP041N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            
                            oMat03.LoadFromDataSource();
                            oMat03.AutoResizeColumns();
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP041H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BaseTime")
                            {
                                oDS_PS_PP041H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "UseMCode")
                            {
                                oForm.Items.Item("UseMName").Specific.Value = dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                                oForm.Items.Item("SMoldNo").Specific.Value = dataHelpClass.GetValue("SELECT U_MoldNo FROM [@PS_PP130H] WHERE U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1); //장비 금형번호
                                oForm.Items.Item("SMoldNm").Specific.Value = dataHelpClass.GetValue("SELECT b.U_Item + '[' + b.U_Callsize +']' FROM [@PS_PP130H] a Inner join [@PS_PP190H] b On a.U_MoldNo = b.Code WHERE a.U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                            }
                            else if (pVal.ItemUID == "SCpCode")
                            {
                                oForm.Items.Item("SCpName").Specific.Value = dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    PS_PP041_OrderInfoLoad();
                                }
                            }
                            else if (pVal.ItemUID == "SMoldNo")
                            {
                                oForm.Items.Item("SMoldNm").Specific.Value = dataHelpClass.GetValue("SELECT U_Item + '[' + U_Callsize +']' FROM [@PS_PP190H] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1); //금형번호 이름
                            }
                            else
                            {
                                oDS_PS_PP041H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }

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
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_PP041_AddMatrixRow01(oMat01.VisualRowCount, false);
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        PS_PP041_AddMatrixRow02(oMat02.VisualRowCount, false);
                        oMat02.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        oMat03.AutoResizeColumns();
                    }

                    PS_PP041_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP041H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP041L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP041M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP041N);
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
                    PS_PP041_FormResize();
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
            string errMessage = string.Empty;
            SAPbouiCOM.BoStatusBarMessageType messageType = BoStatusBarMessageType.smt_Error;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "104" && oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP50101") //멀티의 첫공정일 경우
                        {
                            errMessage = "멀티의 첫공정은 행삭제 할 수 없습니다.";
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "107" && oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP10101") //엔드베어링의 첫공정일 경우
                        {
                            errMessage = "엔드베어링의 첫공정은 행삭제 할 수 없습니다.";
                            BubbleEvent = false;
                            throw new Exception();
                        }

                        if (oLastItemUID01 == "Mat01")
                        {
                            if (PS_PP041_Validate("행삭제01") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            for (i = 1; i <= oMat03.RowCount; i++)
                            {
                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value)
                                {
                                    oDS_PS_PP041N.RemoveRecord(i - 1);
                                    oMat03.DeleteRow(i);
                                    oMat03.FlushToDataSource();
                                    i = 1;
                                    continue;
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

                            oMat01.FlushToDataSource();
                            oDS_PS_PP041L.RemoveRecord(oDS_PS_PP041L.Size - 1);
                            oMat01.LoadFromDataSource();

                            if (oMat01.RowCount == 0)
                            {
                                PS_PP041_AddMatrixRow01(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP041L.GetValue("U_OrdMgNum", oMat01.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP041_AddMatrixRow01(oMat01.RowCount, false);
                                }
                            }
                        }
                        else if (oLastItemUID01 == "Mat02")
                        {
                            for (i = 1; i <= oMat02.VisualRowCount; i++)
                            {
                                oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                            }

                            oMat02.FlushToDataSource();
                            oDS_PS_PP041M.RemoveRecord(oDS_PS_PP041M.Size - 1);
                            oMat02.LoadFromDataSource();

                            if (oMat02.RowCount == 0)
                            {
                                PS_PP041_AddMatrixRow02(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP041M.GetValue("U_WorkCode", oMat02.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP041_AddMatrixRow02(oMat02.RowCount, false);
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
                            if (oDS_PS_PP041N.Size == 1)
                            {
                            }
                            else
                            {
                                oDS_PS_PP041N.RemoveRecord(oDS_PS_PP041N.Size - 1);
                            }
                            oMat03.LoadFromDataSource();
                            
                            for (i = 1; i <= oMat01.RowCount - 1; i++) //공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                            {
                                exist = false;
                                for (j = 1; j <= oMat03.RowCount; j++)
                                {
                                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value)
                                    {
                                        exist = true;
                                    }
                                }
                                
                                if (exist == false) //불량코드테이블에 값이 존재하지 않으면
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP041_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP041_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }

                                    oDS_PS_PP041N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP041N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP041N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OrdMgNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OrdMgNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, messageType);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
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
                        query01 += "                (SELECT MIN(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '20' AND U_OrdGbn IN ('104','107'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '20'";
                        query01 += "            AND U_OrdGbn IN ('104','107')";
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
                        query01 += "                (SELECT MAX(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '20' AND U_OrdGbn IN ('104','107'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '20'";
                        query01 += "            AND U_OrdGbn IN ('104','107')";
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
                    query01 += " WHERE      U_DocType = '20'";
                    query01 += "            AND U_OrdGbn IN ('104','107')";

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
                    query01 += " WHERE      U_DocType = '20'";
                    query01 += "            AND U_OrdGbn IN ('104','107')";

                    oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(query01, 0, 1);
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                }
            }
            catch(Exception ex)
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
                                if (PS_PP041_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PS_PP041_InsertoInventoryGenEntry() == false)
                                {
                                    BubbleEvent = false;
                                    return;
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
                            PS_PP041_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP041_FormItemEnabled();
                            PS_PP041_AddMatrixRow01(0, true);
                            PS_PP041_AddMatrixRow02(0, true);
                            oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
                        if (PS_PP041_FindValidateDocument("@PS_PP040H") == false)
                        {
                            if (PSH_Globals.SBO_Application.Menus.Item("1281").Enabled == true) //찾기메뉴 활성화일때 수행
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("관리자에게 문의바랍니다.");
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
