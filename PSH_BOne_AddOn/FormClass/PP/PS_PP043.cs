using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 포장생산 작업일보등록
    /// </summary>
    internal class PS_PP043 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.DBDataSource oDS_PS_PP043H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP043L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP043M; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP043N; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP043.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP043_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP043");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                PS_PP043_CreateItems();
                PS_PP043_SetComboBox();
                PS_PP043_EnableMenus();
                PS_PP043_SetDocument(oFormDocEntry);
                PS_PP043_FormResize();
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
        private void PS_PP043_CreateItems()
        {
            try
            {
                oDS_PS_PP043H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
                oDS_PS_PP043L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
                oDS_PS_PP043M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
                oDS_PS_PP043N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");

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

                oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");

                oDocType01 = "작업일보등록(공정)";
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
        private void PS_PP043_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "10", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "20", "PSMT지원");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "30", "외주");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "40", "실적");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "50", "일반조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "OrdType", "", "60", "외주조정");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP043", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP043", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP043", "DocType", false);

                oForm.Items.Item("SOrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SOrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('108','109') order by Code", "", false, false);
                oForm.Items.Item("SBPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('108','109') order by Code", "", false, false);
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메뉴 활성화
        /// </summary>
        private void PS_PP043_EnableMenus()
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
        private void PS_PP043_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP043_FormItemEnabled();
                    PS_PP043_AddMatrixRow01(0, true);
                    PS_PP043_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP043_FormItemEnabled();
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
        private void PS_PP043_FormItemEnabled()
        {
            string NextCpInfo;
            int i;
            string query;
            string SuperUserYN = null;
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
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oMat02.Columns.Item("NTime").Editable = true;
                    oForm.Items.Item("Mat03").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;

                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("SOrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("SBPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("SCpCode").Specific.Value = "";

                    PS_PP043_FormClear();

                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else if (oDocType01 == "작업일보등록(공정)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.AddDays(-1).ToString("yyyyMMdd"); //하루 전날
                    oForm.Items.Item("SBPLId").Specific.Select("3", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
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

                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oDS_PS_PP043H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
                    {
                        oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("Mat02").Enabled = false;
                        oForm.Items.Item("Mat03").Enabled = false;
                        oForm.Items.Item("Button01").Enabled = false;
                        oForm.Items.Item("1").Enabled = false;
                    }
                    else
                    {
                        if (oDS_PS_PP043H.GetValue("U_OrdType", 0).ToString().Trim() == "20")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP043H.GetValue("U_OrdType", 0).ToString().Trim() == "30")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP043H.GetValue("U_OrdType", 0).ToString().Trim() == "40")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                        }
                        else if (oDS_PS_PP043H.GetValue("U_OrdType", 0).ToString().Trim() == "10" || oDS_PS_PP043H.GetValue("U_OrdType", 0).ToString().Trim() == "50")
                        {
                            if (oDS_PS_PP043H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104")
                            {

                                query = "  SELECT       PS_PP040H.DocEntry,";
                                query += "              PS_PP040L.LineId,";
                                query += "              CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                                query += "              PS_PP040L.U_OrdGbn AS OrdGbn,";
                                query += "              PS_PP040L.U_PP030HNo AS PP030HNo,";
                                query += "              PS_PP040L.U_PP030MNo AS PP030MNo, ";
                                query += "              PS_PP040L.U_OrdMgNum AS OrdMgNum ";
                                query += " FROM         [@PS_PP040H] PS_PP040H";
                                query += "              LEFT JOIN";
                                query += "              [@PS_PP040L] PS_PP040L";
                                query += "                  ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                                query += " WHERE        PS_PP040H.Canceled = 'N'";
                                query += "              AND PS_PP040L.DocEntry = '" + oDS_PS_PP043H.GetValue("DocEntry", 0) + "'";
                                RecordSet01.DoQuery(query);

                                if (oDS_PS_PP043H.GetValue("DocEntry", 0) != "2")
                                {
                                    SuperUserYN = dataHelpClass.GetValue("select U_UseYN from [@PS_SY001L] a where a.Code ='A007' and a.U_Minor ='PS_PP043' and a.U_RelCd = '" + PSH_Globals.oCompany.UserName + "'", 0, 1);

                                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                    {
                                        if (RecordSet01.Fields.Item("OrdGbn").Value == "104")
                                        {
                                            if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y")
                                            {
                                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE Isnull(PS_PP080L.U_OIGENum,'') = '' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                                {
                                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                                    if (string.IsNullOrEmpty(SuperUserYN))
                                                    {
                                                        oForm.Items.Item("DocEntry").Enabled = false;
                                                        oForm.Items.Item("DocDate").Enabled = false;
                                                        oForm.Items.Item("Mat01").Enabled = false;
                                                        oForm.Items.Item("Mat02").Enabled = false;
                                                        oForm.Items.Item("Mat03").Enabled = false;
                                                        oForm.Items.Item("Button01").Enabled = false;
                                                        oForm.Items.Item("1").Enabled = false;
                                                    }
                                                    else if (SuperUserYN == "Y")
                                                    {
                                                        oForm.Items.Item("DocEntry").Enabled = false;
                                                        oForm.Items.Item("DocDate").Enabled = false;
                                                        oForm.Items.Item("Mat01").Enabled = true;
                                                        oForm.Items.Item("Mat02").Enabled = false;
                                                        oForm.Items.Item("Mat03").Enabled = false;
                                                        oForm.Items.Item("Button01").Enabled = false;
                                                        oForm.Items.Item("1").Enabled = true;
                                                    }

                                                    oMat01.Columns.Item("BQty").Visible = true;
                                                    oMat01.Columns.Item("PSum").Visible = false;
                                                    oMat01.Columns.Item("PWeight").Visible = false;
                                                    oMat01.Columns.Item("YWeight").Visible = false;
                                                    oMat01.Columns.Item("NWeight").Visible = false;
                                                }
                                            }
                                        }

                                        if (RecordSet01.Fields.Item("OrdGbn").Value == "104")
                                        {
                                            NextCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_03 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1);
                                            if (!string.IsNullOrEmpty(NextCpInfo))
                                            {
                                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP040L.U_PP030HNo) + '-' + CONVERT(NVARCHAR,PS_PP040L.U_PP030MNo) = '" + NextCpInfo + "'", 0, 1)) > 0)
                                                {
                                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                                    if (string.IsNullOrEmpty(SuperUserYN))
                                                    {
                                                        oForm.Items.Item("DocEntry").Enabled = false;
                                                        oForm.Items.Item("DocDate").Enabled = false;
                                                        oForm.Items.Item("Mat01").Enabled = false;
                                                        oForm.Items.Item("Mat02").Enabled = false;
                                                        oForm.Items.Item("Mat03").Enabled = false;
                                                        oForm.Items.Item("Button01").Enabled = false;
                                                        oForm.Items.Item("1").Enabled = false;
                                                    }
                                                    else if (SuperUserYN == "Y")
                                                    {
                                                        oForm.Items.Item("DocEntry").Enabled = false;
                                                        oForm.Items.Item("DocDate").Enabled = false;
                                                        oForm.Items.Item("Mat01").Enabled = true;
                                                        oForm.Items.Item("Mat02").Enabled = false;
                                                        oForm.Items.Item("Mat03").Enabled = false;
                                                        oForm.Items.Item("Button01").Enabled = false;
                                                        oForm.Items.Item("1").Enabled = true;
                                                    }

                                                    oMat01.Columns.Item("BQty").Visible = true;
                                                    oMat01.Columns.Item("PSum").Visible = false;
                                                    oMat01.Columns.Item("PWeight").Visible = false;
                                                    oMat01.Columns.Item("YWeight").Visible = false;
                                                    oMat01.Columns.Item("NWeight").Visible = false;
                                                }
                                            }
                                            else
                                            {
                                                //다음공정이 존재하지 않으면 마지막 공정임, 마지막공정일때는 실적등록여부로 적용여부 판정
                                            }
                                        }
                                        RecordSet01.MoveNext();
                                    }
                                }

                                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Items.Item("DocEntry").Enabled = false;
                                oForm.Items.Item("DocDate").Enabled = true;
                                oForm.Items.Item("Mat01").Enabled = true;
                                oForm.Items.Item("Mat02").Enabled = true;
                                oForm.Items.Item("Mat03").Enabled = true;
                                oForm.Items.Item("Button01").Enabled = true;
                                oForm.Items.Item("1").Enabled = true;

                                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                if (string.IsNullOrEmpty(SuperUserYN))
                                {
                                    oForm.Items.Item("DocEntry").Enabled = false;
                                    oForm.Items.Item("DocDate").Enabled = false;
                                    oForm.Items.Item("Mat01").Enabled = false;
                                    oForm.Items.Item("Mat02").Enabled = false;
                                    oForm.Items.Item("Mat03").Enabled = false;
                                    oForm.Items.Item("Button01").Enabled = false;
                                    oForm.Items.Item("1").Enabled = false;
                                }
                                else if (SuperUserYN == "Y")
                                {
                                    oForm.Items.Item("DocEntry").Enabled = false;
                                    oForm.Items.Item("DocDate").Enabled = false;
                                    oForm.Items.Item("Mat01").Enabled = true;
                                    oForm.Items.Item("Mat02").Enabled = false;
                                    oForm.Items.Item("Mat03").Enabled = false;
                                    oForm.Items.Item("Button01").Enabled = false;
                                    oForm.Items.Item("1").Enabled = true;
                                }
                            }
                            else
                            {
                                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Items.Item("DocEntry").Enabled = false;
                                oForm.Items.Item("DocDate").Enabled = true;
                                oForm.Items.Item("Mat01").Enabled = true;
                                oForm.Items.Item("Mat02").Enabled = true;
                                oForm.Items.Item("Mat03").Enabled = true;
                                oForm.Items.Item("Button01").Enabled = true;
                                oForm.Items.Item("1").Enabled = true;
                            }
                            oMat01.Columns.Item("BQty").Visible = true;
                            oMat01.Columns.Item("PSum").Visible = false;
                            oMat01.Columns.Item("PWeight").Visible = false;
                            oMat01.Columns.Item("YWeight").Visible = false;
                            oMat01.Columns.Item("NWeight").Visible = false;
                        }
                        else
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
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
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP043_FormClear()
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
        private void PS_PP043_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP043L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP043L.Offset = oRow;
                oDS_PS_PP043L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP043_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP043M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP043M.Offset = oRow;
                oDS_PS_PP043M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP043_AddMatrixRow03(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP043N.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP043N.Offset = oRow;
                oDS_PS_PP043N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP043_FormResize()
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
        private bool PS_PP043_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            int j;
            double FailQty;
            string errMessage = string.Empty;
            string SuperUserYN;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP043_FormClear();
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

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10" && oForm.Items.Item("OrdType").Specific.Selected.Value != "50")
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
                else if (oMat02.VisualRowCount == 1)
                {
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "107")
                    {
                        oMat02.FlushToDataSource();
                        oDS_PS_PP043M.SetValue("U_WorkCode", 0, "9999999");
                        oDS_PS_PP043M.SetValue("U_WorkName", 0, "조정");
                        oDS_PS_PP043M.SetValue("U_YTime", 0, "1");
                        PS_PP043_AddMatrixRow02(1, false);
                        oMat02.LoadFromDataSource();
                    }
                    else
                    {
                        errMessage = "작업자정보 라인이 존재하지 않습니다.";
                        oMat02.SelectRow(oMat02.VisualRowCount, true, false);
                        throw new Exception();
                    }
                }
                else if (oMat03.VisualRowCount == 0)
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
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 && string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
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

                //비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
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
                //비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E

                //불량정보 입력 체크_S (2012.03.20 송명규 추가)                
                SuperUserYN = dataHelpClass.GetValue("select U_UseYN from [@PS_SY001L] a where a.Code ='A007' and a.U_Minor = 'PS_PP043' and a.U_RelCd = '" + PSH_Globals.oCompany.UserName + "'", 0, 1);

                //슈퍼유저가 아니면 불량정보 필수 입력
                if (string.IsNullOrEmpty(SuperUserYN))
                {
                    for (i = 1; i <= oMat03.VisualRowCount - 1; i++)
                    {
                        //해당 작업지시의 재작업 여부 조회
                        if (dataHelpClass.GetValue("SELECT U_ReWorkYN FROM [@PS_PP030M] WHERE Convert(Nvarchar(50),DocEntry) + '-' + Convert(Nvarchar(20), U_LineId) = '" + oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                        {
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(i).Specific.Value)) //불량코드
                            {
                                errMessage = "재작업 시 불량정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat03.Columns.Item("CsCpCode").Cells.Item(i).Specific.Value)) //원인공정코드
                            {
                                errMessage = "재작업 시 원인공정정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("CsCpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat03.Columns.Item("CsWkCode").Cells.Item(i).Specific.Value)) //작업자코드
                            {
                                errMessage = "재작업 시 작업자정보는 필수입니다. 확인하십시오.";
                                oMat03.Columns.Item("CsWkCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }
                }
                else
                {
                }
                //불량정보 입력 체크_E (2012.03.20 송명규 추가)

                if (PS_PP043_Validate("검사01") == false)
                {
                    errMessage = " ";
                    throw new Exception();
                }

                oDS_PS_PP043L.RemoveRecord(oDS_PS_PP043L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP043M.RemoveRecord(oDS_PS_PP043M.Size - 1);
                oMat02.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
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
        private bool PS_PP043_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 일반,조정인경우
                {
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20"
                      || oForm.Items.Item("OrdType").Specific.Selected.Value == "30"
                      || oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 PSMT지원(20), 외주(30), 실적(40)인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();
                }

                if (ValidateType == "검사01")
                {
                    for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                        {
                            errMessage = "제품코드가 존재하지 않습니다.";
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
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
                else if (ValidateType == "행삭제01")
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM095H] a Inner JOIN [@PS_MM095L] b ON a.DocEntry = b.DocEntry WHERE a.Canceled = 'N' AND b.U_PP040Doc = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND b.U_PP040Lin = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                        {
                            errMessage = "원재료 불출된 행입니다. 행 삭제할 수 없습니다. 원재료 불출 취소 후 행 삭제 바랍니다.";
                            throw new Exception();
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
                else if (ValidateType == "수정01")
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
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
                        errMessage = "이미 취소된 문서 입니다. 취소할 수 없습니다.";
                        throw new Exception();
                    }

                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
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

            return returnValue;
        }

        /// <summary>
        /// OrderInfoLoad
        /// </summary>
        private void PS_PP043_OrderInfoLoad()
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
                        else //값이 선택되었다면 선택된 값
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
                        PS_PP043_AddMatrixRow01(0, true);
                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP043_AddMatrixRow02(0, true);
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
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FindValidateDocument : 포장생산 작업일보등록 문서인지 조회
        /// </summary>
        /// <param name="ObjectType"></param>
        /// <returns></returns>
        private bool PS_PP043_FindValidateDocument(string ObjectType)
        {
            bool returnValue = false;
            string query;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                query = "  SELECT   DocEntry";
                query += " FROM     [" + ObjectType + "]";
                query += " WHERE    DocEntry = " + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                if (oDocType01 == "작업일보등록(작지)")
                {
                    query += " AND U_DocType = '10'";
                }
                else if (oDocType01 == "작업일보등록(공정)")
                {
                    query += " AND U_DocType = '20'";
                }
                RecordSet01.DoQuery(query);

                if (RecordSet01.RecordCount == 0)
                {
                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        errMessage = "작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.";
                    }
                    else if ((oDocType01 == "작업일보등록(공정)"))
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
        /// 출고 DI(미사용 메소드인 것으로 추정, 사용확인 후 삭제 필요)
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        //private bool PS_PP043_InsertoInventoryGenExit(short ChkType)
        //{
        //    bool returnValue = false;
        //    int RetVal;
        //    string sQry;
        //    string errCode = string.Empty;
        //    string errDIMsg = string.Empty;
        //    int errDICode = 0;
        //    string afterDIDocNum;
        //    int i;
        //    int oRow;
        //    int Cnt = 0;
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        //    SAPbobsCOM.Documents DI_oInventoryGenExit = null;
        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    try
        //    {
        //        PSH_Globals.oCompany.StartTransaction();

        //        //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
        //        if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
        //        {
        //            errCode = "2";
        //            throw new Exception();
        //        }

        //        for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //        {
        //            if (oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107")
        //            {
        //                //107010002(END BEARING #44),107010004(END BEARING #2) 일경우에는 원자재가 스크랩을 이용이므로, 불출될 원자재가 없음.
        //                if (oDS_PS_PP043L.GetValue("U_ItemCode", oRow).ToString().Trim() != "107010002" && oDS_PS_PP043L.GetValue("U_ItemCode", oRow).ToString().Trim() != "107010004")
        //                {
        //                    if (oDS_PS_PP043L.GetValue("U_Sequence", oRow).ToString() == "1") //첫공정일 경우
        //                    {
        //                        sQry = "  select    b.docentry";
        //                        sQry += " from      [@PS_PP040L] a";
        //                        sQry += "           inner join";
        //                        sQry += "           [@PS_PP040H] b";
        //                        sQry += "               on a.docentry=b.docentry ";
        //                        sQry += " where     a.U_OrdGbn in ('104','107')";
        //                        sQry += "           and b.canceled <> 'Y' ";
        //                        sQry += "           and a.U_PP030HNo = '" + oDS_PS_PP043L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
        //                        sQry += "           and a.U_Sequence = '" + oDS_PS_PP043L.GetValue("U_Sequence", oRow).ToString().Trim() + "'";
        //                        oRecordSet.DoQuery(sQry);

        //                        //처음 작업일보 등록시
        //                        if (oRecordSet.RecordCount < 1)
        //                        {
        //                            Cnt += 1;
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        if (Cnt < 1) //출고 DI API 실행 조건에 해당하는 Data가 없을 때
        //        { 
        //            returnValue = true;
        //            return returnValue;
        //        }

        //        DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

        //        i = 1;
        //        //Header
        //        DI_oInventoryGenExit.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
        //        DI_oInventoryGenExit.TaxDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
        //        DI_oInventoryGenExit.Comments = "작업일보등록(" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + ") 출고";

        //        //Line
        //        for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //        {
        //            if (oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107") //멀티 & 엔드베어링일경우
        //            {
        //                if (oDS_PS_PP043L.GetValue("U_Sequence", oRow).ToString().Trim() == "1") //첫공정일 경우
        //                {
        //                    sQry = "  select    b.docentry";
        //                    sQry += " from      [@PS_PP040L] a";
        //                    sQry += "           inner join";
        //                    sQry += "           [@PS_PP040H] b";
        //                    sQry += "               on a.docentry = b.docentry ";
        //                    sQry += " where     a.U_OrdGbn in ('104','107')";
        //                    sQry += "           and b.canceled <> 'Y' ";
        //                    sQry += "           and a.U_PP030HNo = '" + oDS_PS_PP043L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
        //                    sQry += "           and a.U_Sequence = '" + oDS_PS_PP043L.GetValue("U_Sequence", oRow).ToString().Trim() + "'";
        //                    oRecordSet.DoQuery(sQry);

        //                    if (oRecordSet.RecordCount < 1) //처음 작업일보 등록시
        //                    {
        //                        if (DI_oInventoryGenExit.Lines.Count < i)
        //                        {
        //                            DI_oInventoryGenExit.Lines.Add();
        //                            DI_oInventoryGenExit.Lines.BatchNumbers.Add();
        //                        }

        //                        sQry = "select U_ItemCode, U_ItemName, U_BatchNum, U_Weight from [@PS_PP030L] where docentry = '" + oDS_PS_PP043L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
        //                        oRecordSet.DoQuery(sQry);

        //                        DI_oInventoryGenExit.Lines.SetCurrentLine(i - 1);
        //                        DI_oInventoryGenExit.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
        //                        DI_oInventoryGenExit.Lines.ItemDescription = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
        //                        DI_oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim();
        //                        DI_oInventoryGenExit.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
        //                        DI_oInventoryGenExit.Lines.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());

        //                        sQry = "select TOP 1 WhsCode from [OIBT] where Quantity <> 0 ";
        //                        sQry = sQry + "and ItemCode = '" + oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim() + "' ";
        //                        sQry = sQry + "and BatchNum = '" + oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim() + "'";
        //                        oRecordSet.DoQuery(sQry);
        //                        DI_oInventoryGenExit.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim();

        //                        i += 1;
        //                    }
        //                }
        //            }
        //        }

        //        RetVal = DI_oInventoryGenExit.Add();
        //        if (0 != RetVal)
        //        {
        //            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
        //            errCode = "1";
        //            throw new Exception();
        //        }
        //        if (ChkType != 2)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        else
        //        {
        //            PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);

        //            i = 1;
        //            for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //            {
        //                if (oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "104" || oDS_PS_PP043L.GetValue("U_OrdGbn", oRow).ToString().Trim() == "107") //멀티 & 엔드베어링일경우
        //                {
        //                    if (oDS_PS_PP043L.GetValue("U_Sequence", oRow).ToString().Trim() == "1") //첫공정일 경우
        //                    {
        //                        oDS_PS_PP043L.SetValue("U_OutDoc", oRow, afterDIDocNum);
        //                        oDS_PS_PP043L.SetValue("U_OutLin", oRow, Convert.ToString(i));

        //                        i += 1;
        //                    }
        //                }
        //            }
        //            oMat01.LoadFromDataSource();

        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        }

        //        returnValue = true;
        //    }
        //    catch(Exception ex)
        //    {
        //        if (PSH_Globals.oCompany.InTransaction)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }

        //        if (errCode == "1")
        //        {
        //            PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
        //        }
        //        else if (errCode == "2")
        //        {
        //            PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //        if (DI_oInventoryGenExit != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
        //        }
        //    }

        //    return returnValue;
        //}

        /// <summary>
        /// 입고DI(출고 취소)(미사용 메소드인 것으로 추정, 사용확인 후 삭제 필요)
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        //private bool PS_PP043_InsertoInventoryGenEntry(short ChkType)
        //{
        //    bool returnValue = false;
        //    int RetVal;
        //    string sQry;
        //    int i;
        //    int oRow;
        //    int Cnt = 0;
        //    string errCode = string.Empty;
        //    string errDIMsg = string.Empty;
        //    int errDICode = 0;
        //    string afterDIDocNum;

        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        //    SAPbobsCOM.Documents DI_oInventoryGenEntry = null;
        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    try
        //    {
        //        PSH_Globals.oCompany.StartTransaction();

        //        //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
        //        if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
        //        {
        //            errCode = "2";
        //            throw new Exception();
        //        }

        //        for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //        {
        //            if (!string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_OutDoc", oRow).ToString().Trim())) //출고 문서가 있는경우
        //            {
        //                Cnt += 1;
        //            }
        //        }

        //        if (Cnt < 1) //DI API 실행 조건에 해당하는 Data가 없을 때
        //        {
        //            returnValue = true;
        //            return returnValue;
        //        }

        //        DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

        //        i = 1;

        //        //Header
        //        DI_oInventoryGenEntry.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
        //        DI_oInventoryGenEntry.TaxDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
        //        DI_oInventoryGenEntry.Comments = "작업일보등록(" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + ") 출고 취소";

        //        //Line
        //        for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //        {
        //            if (!string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_OutDoc", oRow).ToString().Trim())) //출고 문서가 있는경우
        //            {
        //                if (DI_oInventoryGenEntry.Lines.Count < i)
        //                {
        //                    DI_oInventoryGenEntry.Lines.Add();
        //                    DI_oInventoryGenEntry.Lines.BatchNumbers.Add();
        //                }

        //                sQry = "select U_ItemCode,U_ItemName,U_BatchNum,U_Weight from [@PS_PP030L] where docentry = '" + oDS_PS_PP043L.GetValue("U_PP030HNo", oRow).ToString().Trim() + "'";
        //                oRecordSet.DoQuery(sQry);

        //                DI_oInventoryGenEntry.Lines.SetCurrentLine(i - 1);
        //                DI_oInventoryGenEntry.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
        //                DI_oInventoryGenEntry.Lines.ItemDescription = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
        //                DI_oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRecordSet.Fields.Item("U_BatchNum").Value.ToString().Trim();
        //                DI_oInventoryGenEntry.Lines.BatchNumbers.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
        //                DI_oInventoryGenEntry.Lines.Quantity = Convert.ToDouble(oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());

        //                //출고된 창고 select
        //                sQry = "select WhsCode from [IGE1]";
        //                sQry = sQry + "where docentry = '" + oDS_PS_PP043L.GetValue("U_OutDoc", oRow).ToString().Trim() + "' ";
        //                sQry = sQry + "and linenum = '" + Convert.ToString(Convert.ToInt32(oDS_PS_PP043L.GetValue("U_OutLin", oRow).ToString().Trim()) - 1) + "'";
        //                oRecordSet.DoQuery(sQry);
        //                DI_oInventoryGenEntry.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value.ToString().Trim();

        //                i += 1;
        //            }
        //        }

        //        RetVal = DI_oInventoryGenEntry.Add();
        //        if (0 != RetVal)
        //        {
        //            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
        //            errCode = "1";
        //            throw new Exception();
        //        }
        //        if (ChkType != 2)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        else
        //        {
        //            PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);

        //            i = 1;
        //            for (oRow = 0; oRow <= oDS_PS_PP043L.Size - 1; oRow++)
        //            {
        //                //출고되어진 문서가 있는경우
        //                if (!string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_OutDoc", oRow).ToString().Trim()))
        //                {
        //                    //update
        //                    sQry = "Update [@PS_PP040L] set U_OutDocC = '" + afterDIDocNum + "'";
        //                    sQry = sQry + ", U_OutLinC = '" + i + "' ";
        //                    sQry = sQry + "where docentry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' ";
        //                    sQry = sQry + "and visorder = '" + oRow + "'";
        //                    oRecordSet.DoQuery(sQry);

        //                    i += 1;
        //                }
        //            }
        //            oMat01.LoadFromDataSource();

        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        }

        //        returnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (PSH_Globals.oCompany.InTransaction)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }

        //        if (errCode == "1")
        //        {
        //            PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
        //        }
        //        else if (errCode == "2")
        //        {
        //            PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //        if (DI_oInventoryGenEntry != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
        //        }
        //    }

        //    return returnValue;
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
            int i;
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP043_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //(미사용 메소드인 것으로 추정, 사용확인 후 삭제 필요)
                            //if (PS_PP043_InsertoInventoryGenExit(2) == false)
                            //{
                            //    BubbleEvent = false;
                            //    return;
                            //}
                            //else
                            //{
                            //}

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP043_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value;
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "2") //취소버튼 누를시 저장할 자료가 있으면 메시지 표시
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oMat01.VisualRowCount > 1)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소") == 2)
                                {
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP043_OrderInfoLoad();
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
                            for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                            {
                                totTime += Convert.ToDouble(oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value);
                            }

                            if (totTime > 0)
                            {
                                unitTime = Convert.ToDouble((totTime / (oMat01.VisualRowCount - 1)).ToString("#,##0.##"));
                                unitRemainTime = Convert.ToDouble((totTime - unitTime * (oMat01.VisualRowCount - 1)).ToString("#,##0.##"));
                                
                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    if (i != oMat01.VisualRowCount - 2)
                                    {
                                        oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                    }
                                    else
                                    {
                                        oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
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
                                PS_PP043_FormItemEnabled();
                                PS_PP043_AddMatrixRow02(0, true);
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
                                    PS_PP043_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP043_FormItemEnabled();
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
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "108" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "109")
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value == "")
                                        {
                                            PS_SM010 tempForm = new PS_SM010();
                                            tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
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
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "WorkCode")
                        {
                            if ((oForm.Items.Item("BaseTime").Specific.Value == "" ? 0 : Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value)) == 0)
                            {
                                errMessage = "기준시간을 입력하지 않았습니다.";
                                messageType = BoStatusBarMessageType.smt_Warning;
                                oForm.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                BubbleEvent = false;
                                throw new Exception();
                            }
                        }
                    }
                    if (pVal.ItemUID == "SCpCode")
                    {
                        if (oForm.Items.Item("SOrdGbn").Specific.Selected.Value == "선택")
                        {
                            errMessage = "작업구분이 선택되지 않았습니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                    }
                    if (pVal.ItemUID == "SMoldNo")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("SCpCode").Specific.Value))
                        {
                            errMessage = "공정이 선택되지 않았습니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        if (string.IsNullOrEmpty(oForm.Items.Item("SMoldNo").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    if (pVal.ItemUID == "UseMCode")
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
                                oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP043_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP043M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP043_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                            }
                            else
                            {
                                oDS_PS_PP043N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "특정컬럼")
                            {
                                oDS_PS_PP043H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if ((pVal.ItemUID == "SBPLId" || pVal.ItemUID == "SOrdGbn"))
                            {
                            }
                            else
                            {
                                oDS_PS_PP043H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
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
                    if (pVal.ItemUID == "Opt02")
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
                    if (pVal.ItemUID == "Opt03")
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
                                        PS_PP043_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP043_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }
                                    oDS_PS_PP043N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP043N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP043N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string query;
            double unitTime;
            double unitRemainTime;
            double time;
            double hour;
            double minute;
            string BPLId;
            string CpCode;
            string CpName;
            string OrdGbn;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
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
                            if (PS_PP043_Validate("수정01") == false)
                            {
                                oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrdMgNum")
                                {
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택") //작업구분에 값이 없으면 작업지시가 불러오기전
                                    {
                                        oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                    }
                                    else //작업지시가 선택된상태
                                    {
                                        if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 일반,조정
                                        {
                                            BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                                            CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
                                            CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
                                            OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();

                                            for (i = 1; i <= oMat01.RowCount; i++)
                                            {
                                                //현재 입력한 값이 이미 입력되어 있는경우
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row)
                                                {
                                                    errMessage = "이미 입력한 공정입니다.";
                                                    oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                                    {
                                                        PS_PP043_AddMatrixRow01(pVal.Row, false);
                                                    }
                                                    throw new Exception();
                                                }
                                            }
                                            
                                            query = "Select a.ItemCode, a.ItemName From OITM a Where a.ItmsGrpCod = '102' ";
                                            query += "And Not Exists (Select * from [@PS_MM002H] b Where a.ItemCode = b.U_ItemCode And b.U_Type = '2')  ";
                                            query += "And a.ItemCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                            RecordSet01.DoQuery(query);

                                            if (RecordSet01.RecordCount == 0)
                                            {
                                                errMessage = "제품코드가 아니거나 없는 원재료 판매 제품입니다.";
                                                oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                throw new Exception();
                                            }
                                            else
                                            {
                                                oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP043L.SetValue("U_Sequence", pVal.Row - 1, codeHelpClass.Right(CpCode, 1));
                                                oDS_PS_PP043L.SetValue("U_CpCode", pVal.Row - 1, CpCode);
                                                oDS_PS_PP043L.SetValue("U_CpName", pVal.Row - 1, CpName);
                                                oDS_PS_PP043L.SetValue("U_OrdGbn", pVal.Row - 1, OrdGbn);
                                                oDS_PS_PP043L.SetValue("U_BPLId", pVal.Row - 1, BPLId);
                                                oDS_PS_PP043L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP043L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                                oDS_PS_PP043L.SetValue("U_PSum", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_BQty", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_LineId", pVal.Row - 1, "");
                                                oDS_PS_PP043L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                oDS_PS_PP043L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP043L.SetValue("U_PP030MNo", pVal.Row - 1, codeHelpClass.Right(CpCode, 1));

                                                //불량코드테이블
                                                if (oMat03.VisualRowCount == 0)
                                                {
                                                    PS_PP043_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    PS_PP043_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                }
                                                oDS_PS_PP043N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP043N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, CpCode);
                                                oDS_PS_PP043N.SetValue("U_CpName", oMat03.VisualRowCount - 1, CpName);
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

                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP043_AddMatrixRow01(pVal.Row, false);
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "PQty")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                        oDS_PS_PP043L.SetValue("U_YQty", pVal.Row - 1, "0");
                                        oDS_PS_PP043L.SetValue("U_NQty", pVal.Row - 1, oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                        oDS_PS_PP043L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                    }
                                    oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP043L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                                else if (pVal.ColUID == "NQty")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP043L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                    }
                                    else
                                    {
                                        oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                        oDS_PS_PP043L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                }
                                else if (pVal.ColUID == "WorkTime")
                                {
                                    oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                }
                                else if (pVal.ColUID == "MachCode")
                                {
                                    oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                    oDS_PS_PP043L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP043L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "WorkCode")
                            {
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                oDS_PS_PP043M.SetValue("U_WorkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                oDS_PS_PP043M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value); //기준시간을 작업시간에 입력
                                oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP043M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP043_AddMatrixRow02(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "NStart")
                            {
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP043M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP043M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount - 1);
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
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
                                    oDS_PS_PP043M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP043M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount - 1);
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
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
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP043M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP043M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount - 1);
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
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
                                    oDS_PS_PP043M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP043M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티
                                    {
                                        if (oMat02.VisualRowCount > 1)
                                        {
                                            if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                            {
                                                //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                                if (Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) > 0)
                                                {
                                                    unitTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) / (oMat01.VisualRowCount - 1);
                                                    unitRemainTime = Convert.ToDouble(oDS_PS_PP043M.GetValue("U_YTime", pVal.Row - 1)) - (unitTime * (oMat01.VisualRowCount - 1));
                                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                    {
                                                        if (i != oMat01.VisualRowCount - 2)
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
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
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                oDS_PS_PP043M.SetValue("U_TTime", pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
                                
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //멀티
                                {
                                    if (oMat02.VisualRowCount > 1)
                                    {
                                        if (dataHelpClass.GetValue("SELECT TOP 1 U_CpCode FROM [@PS_PP030M] WHERE DocEntry = '" + oMat01.Columns.Item("PP030HNo").Cells.Item(1).Specific.Value + "' ORDER BY U_Sequence ASC", 0, 1) != oMat01.Columns.Item("CpCode").Cells.Item(1).Specific.Value)
                                        {
                                            //해당작지의 첫공정과 공정정보의 공정이 다르면 분배 '//V_MILL일때만 해당.. 엔드베어링에서는 어떻게 동작하는지 정의필요
                                            if (Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0)
                                            {
                                                unitTime = Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / (oMat01.VisualRowCount - 1);
                                                unitRemainTime = Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) - (unitTime * (oMat01.VisualRowCount - 1));
                                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                                {
                                                    if (i != oMat01.VisualRowCount - 2)
                                                    {
                                                        oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime));
                                                    }
                                                    else
                                                    {
                                                        oDS_PS_PP043L.SetValue("U_WorkTime", i, Convert.ToString(unitTime + unitRemainTime));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                oDS_PS_PP043M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP043N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP043N.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ColUID == "CsCpCode")
                            {
                                oDS_PS_PP043N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP043N.SetValue("U_CsCpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ColUID == "CsWkCode")
                            {
                                oDS_PS_PP043N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP043N.SetValue("U_CsWkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT T0.lastName+T0.firstName FROM OHEM T0 WHERE T0.U_MSTCOD = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP043N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP043H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BaseTime")
                            {
                                oDS_PS_PP043H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "UseMCode")
                            {
                                oForm.Items.Item("UseMName").Specific.Value = dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                                oForm.Items.Item("SMoldNo").Specific.Value = dataHelpClass.GetValue("SELECT U_MoldNo FROM [@PS_PP130H] WHERE U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                                oForm.Items.Item("SMoldNm").Specific.Value = dataHelpClass.GetValue("SELECT b.U_Item + '[' + b.U_Callsize +']' FROM [@PS_PP130H] a Inner join [@PS_PP190H] b On a.U_MoldNo = b.Code WHERE a.U_MachCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                            }
                            else if (pVal.ItemUID == "SCpCode")
                            {
                                oForm.Items.Item("SCpName").Specific.Value = dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    PS_PP043_OrderInfoLoad();
                                }
                            }
                            else if (pVal.ItemUID == "SMoldNo")
                            {
                                oForm.Items.Item("SMoldNm").Specific.Value = dataHelpClass.GetValue("SELECT U_Item + '[' + U_Callsize +']' FROM [@PS_PP190H] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                            }
                            else
                            {
                                oDS_PS_PP043H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
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
                oForm.Freeze(false) ;
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
                    PS_PP043_FormItemEnabled();
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_PP043_AddMatrixRow01(oMat01.VisualRowCount, false);
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        PS_PP043_AddMatrixRow02(oMat02.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP043H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP043L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP043M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP043N);
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
                    PS_PP043_FormResize();
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
            bool Exist;
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
                        else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "107" && oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP10101")
                        {
                            errMessage = "엔드베어링의 첫공정은 행삭제 할 수 없습니다.";
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        
                        if (oLastItemUID01 == "Mat01")
                        {
                            if (PS_PP043_Validate("행삭제01") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        
                            for (i = 1; i <= oMat03.RowCount; i++)
                            {
                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value)
                                {
                                    oDS_PS_PP043N.RemoveRecord(i - 1);
                                    oMat03.DeleteRow(i);
                                    oMat03.FlushToDataSource();
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
                            oDS_PS_PP043L.RemoveRecord(oDS_PS_PP043L.Size - 1);
                            oMat01.LoadFromDataSource();
                            if (oMat01.RowCount == 0)
                            {
                                PS_PP043_AddMatrixRow01(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP043L.GetValue("U_OrdMgNum", oMat01.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP043_AddMatrixRow01(oMat01.RowCount, false);
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
                            oDS_PS_PP043M.RemoveRecord(oDS_PS_PP043M.Size - 1);
                            oMat02.LoadFromDataSource();
                            if (oMat02.RowCount == 0)
                            {
                                PS_PP043_AddMatrixRow02(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP043M.GetValue("U_WorkCode", oMat02.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP043_AddMatrixRow02(oMat02.RowCount, false);
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

                            if (oDS_PS_PP043N.Size == 1) //사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
                            {
                            }
                            else
                            {
                                oDS_PS_PP043N.RemoveRecord(oDS_PS_PP043N.Size - 1);
                            }
                            oMat03.LoadFromDataSource();
                            
                            for (i = 1; i <= oMat01.RowCount - 1; i++) //공정 테이블에는 있는데 불량 테이블에 존재하지 않는 값이 있는경우 불량테이블에 값을 추가함
                            {
                                Exist = false;
                                for (j = 1; j <= oMat03.RowCount; j++)
                                {
                                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value)
                                    {
                                        Exist = true;
                                    }
                                }
                                
                                if (Exist == false) //불량코드테이블에 값이 존재하지 않으면
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP043_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP043_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }
                                    oDS_PS_PP043N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP043N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP043N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
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
                        query01 += "                (SELECT MIN(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '20' AND U_OrdGbn IN ('108','109'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '20'";
                        query01 += "            AND U_OrdGbn IN ('108','109')";
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
                        query01 += "                (SELECT MAX(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '20' AND U_OrdGbn IN ('108','109'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '20'";
                        query01 += "            AND U_OrdGbn IN ('108','109')";
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
                    query01 += "            AND U_OrdGbn IN ('108','109')";

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
                    query01 += "            AND U_OrdGbn IN ('108','109')";

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
                                if (PS_PP043_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                //(미사용 메소드인 것으로 추정, 사용확인 후 삭제 필요)
                                //if (PS_PP043_InsertoInventoryGenEntry(2) == false)
                                //{
                                //    BubbleEvent = false;
                                //    return;
                                //}
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 취소할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286"://닫기
                            break;
                        case "1293"://행삭제
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
                            PS_PP043_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP043_FormItemEnabled();
                            PS_PP043_AddMatrixRow01(0, true);
                            PS_PP043_AddMatrixRow02(0, true);
                            oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
                        if (PS_PP043_FindValidateDocument("@PS_PP040H") == false)
                        {
                            //찾기메뉴 활성화일때 수행
                            if (PSH_Globals.SBO_Application.Menus.Item("1281").Enabled == true)
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
