using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작업일보등록(방산부품)
    /// </summary>
    internal class PS_PP044 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.Matrix oMat04;
        private SAPbouiCOM.Matrix oMat05;
        private SAPbouiCOM.DBDataSource oDS_PS_PP044H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP044L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP044M; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP044N; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP044T; //사원조회라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP044U; //사원조회라인
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
        private string oGubun;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP044.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP044_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP044");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                PS_PP044_CreateItems();
                PS_PP044_ComboBox_Setting();
                PS_PP044_CF_ChooseFromList();
                PS_PP044_EnableMenus();
                PS_PP044_SetDocument(oFormDocEntry);
                PS_PP044_FormResize();
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
        private void PS_PP044_CreateItems()
        {
            try
            {
                oDS_PS_PP044H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
                oDS_PS_PP044L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
                oDS_PS_PP044M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
                oDS_PS_PP044N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");
                oDS_PS_PP044T = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_PP044U = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oMat04 = oForm.Items.Item("Mat04").Specific;
                oMat04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat04.AutoResizeColumns();

                oMat05 = oForm.Items.Item("Mat05").Specific;
                oMat05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat05.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Gubun").Specific.DataBind.SetBound(true, "", "Gubun");

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
        private void PS_PP044_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "10", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "20", "PSMT지원");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "30", "외주");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "40", "실적");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "50", "일반조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "60", "외주조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "70", "설계시간");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP044", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP044", "DocType", false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE IN('102','602') order by Code", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");

                oForm.Items.Item("Gubun").Specific.ValidValues.Add("선택", "선택");
                oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작업구분코드(2014.04.15 송명규 수정)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'P203'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";

                if (oMat01.Columns.Item("WorkCls").ValidValues.Count > 0)
                {
                    for (int loopCount = 0; loopCount <= oMat01.Columns.Item("WorkCls").ValidValues.Count - 1; loopCount++)
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
        private void PS_PP044_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;

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
        /// 메뉴설정
        /// </summary>
        private void PS_PP044_EnableMenus()
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
        private void PS_PP044_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP044_FormItemEnabled();
                    PS_PP044_AddMatrixRow01(0, true);
                    PS_PP044_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP044_FormItemEnabled();
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
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_PP044_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oForm.Items.Item("Mat03").Enabled = true;
                    oMat02.Columns.Item("NTime").Editable = true; //비가동시간만 사용

                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    if (string.IsNullOrEmpty(oOrdGbn))
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select("102", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    if (oGubun == "선택" || string.IsNullOrEmpty(oGubun))
                    {
                        oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        oForm.Items.Item("Gubun").Specific.Select(oGubun, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    PS_PP044_FormClear();

                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        oDS_PS_PP044H.SetValue("U_DocType", 0, "10");
                    }
                    else if (oDocType01 == "작업일보등록(공정)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    if (string.IsNullOrEmpty(oDocdate))
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

                    if (oGubun == "선택" || string.IsNullOrEmpty(oGubun))
                    {
                        oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        oForm.Items.Item("Gubun").Specific.Select(oGubun, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oDS_PS_PP044H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
                    {
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
                        if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "10"
                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "50"
                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "60"
                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "70") //조정, 설계
                        {
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
                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "20") //PSMT
                        {
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
                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "30") //외주
                        {
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
                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "40") //실적
                        {
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
        private void PS_PP044_FormClear()
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
        private void PS_PP044_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP044L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP044L.Offset = oRow;
                oDS_PS_PP044L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP044L.SetValue("U_WorkCls", oRow, "A"); //작업구분을 기본으로 선택(2014.04.15 송명규 추가)
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
        private void PS_PP044_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP044M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP044M.Offset = oRow;
                oDS_PS_PP044M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP044_AddMatrixRow03(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP044N.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP044N.Offset = oRow;
                oDS_PS_PP044N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP044_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat02").Top = (oForm.Height / 5) * 2;
                oForm.Items.Item("Mat02").Left = 7;
                oForm.Items.Item("Mat02").Height = ((oForm.Height - 170) / 3 * 1) - 20;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 14;

                oForm.Items.Item("Mat03").Top = (oForm.Height / 5) * 2;
                oForm.Items.Item("Mat03").Left = oForm.Width / 2;
                oForm.Items.Item("Mat03").Height = ((oForm.Height - 170) / 3 * 1) - 20;
                oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 14;

                oForm.Items.Item("Mat04").Top = 10;
                oForm.Items.Item("Mat04").Left = oForm.Width / 2;
                oForm.Items.Item("Mat04").Height = (oForm.Height / 4) - 10;
                oForm.Items.Item("Mat04").Width = oForm.Width / 3 - 25;
                oMat04.AutoResizeColumns();

                oForm.Items.Item("Mat05").Top = oForm.Items.Item("Mat04").Top + oForm.Items.Item("Mat04").Height;
                oForm.Items.Item("Mat05").Left = oForm.Width / 2;
                oForm.Items.Item("Mat05").Height = oForm.Items.Item("Mat04").Height / 2;
                oForm.Items.Item("Mat05").Width = oForm.Width / 3 - 25;
                oMat04.AutoResizeColumns();

                oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;
                oForm.Items.Item("Mat01").Left = 7;
                oForm.Items.Item("Mat01").Height = oForm.Height / 4;
                oForm.Items.Item("Mat01").Width = oForm.Width - 21;

                oForm.Items.Item("Opt01").Top = oForm.Items.Item("Mat02").Top - 20;
                oForm.Items.Item("Opt02").Top = oForm.Items.Item("Mat03").Top - 20;
                oForm.Items.Item("EmpChk").Top = oForm.Items.Item("Mat02").Top - 20;
                oForm.Items.Item("Button03").Top = oForm.Items.Item("Mat02").Top - 20;

                oForm.Items.Item("Opt01").Left = 10;
                oForm.Items.Item("Opt02").Left = oForm.Width / 2;
                oForm.Items.Item("Opt03").Left = 10;
                oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 5;

                oForm.Items.Item("Button03").Left = oForm.Items.Item("Mat02").Width - oForm.Items.Item("Button03").Width;

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
        private bool PS_PP044_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            int j;
            double FailQty;
            double sYTime = 0;
            double sNTime = 0;
            double sTTime = 0;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (Convert.ToInt32(dataHelpClass.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.Value + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'", 0, 1)) > 0)
                {
                    errMessage = "해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.";
                    throw new Exception();
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP044_FormClear();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10"
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "20"
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "50"
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "60"
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "70")
                {
                    errMessage = "작업타입이 일반, PSMT지원, 조정, 설계가 아닙니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                {
                    errMessage = "작지번호는 필수입니다.";
                    oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "공정정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                if (oMat02.VisualRowCount == 1)
                {
                    errMessage = "작업자정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (j = 1; j <= oMat02.VisualRowCount; j++)
                {
                    sYTime += Convert.ToDouble(oMat02.Columns.Item("YTime").Cells.Item(j).Specific.Value == "" ? "0" : oMat02.Columns.Item("YTime").Cells.Item(j).Specific.Value);
                    sNTime += Convert.ToDouble(oMat02.Columns.Item("NTime").Cells.Item(j).Specific.Value == "" ? "0" : oMat02.Columns.Item("NTime").Cells.Item(j).Specific.Value);
                    sTTime += Convert.ToDouble(oMat02.Columns.Item("TTime").Cells.Item(j).Specific.Value == "" ? "0" : oMat02.Columns.Item("TTime").Cells.Item(j).Specific.Value);
                }

                if (oMat03.VisualRowCount == 0)
                {
                    errMessage = "불량정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                //마감상태 체크_S(2017.11.23 송명규 추가)
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                //마감상태 체크_E(2017.11.23 송명규 추가)

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작지문서번호는 필수입니다.";
                        oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50"
                     && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        if (sYTime > 0) //작업시간이 0 보다 클때
                        {
                            if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                            {
                                errMessage = "생산수량은 필수입니다.";
                                oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50"
                     && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60"
                     && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "70")
                    {
                        if (sYTime > 0) //작업시간이 0 보다 클때
                        {
                            if (Convert.ToDouble(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)
                            {
                                errMessage = "실동시간은 필수입니다.";
                                oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }

                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105"
                     || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "106") //기계공구, 몰드일 경우만 작업완료여부 필수 체크
                    {
                        if (oMat01.Columns.Item("CompltYN").Cells.Item(i).Specific.Value == "%")
                        {
                            errMessage = "작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오.";
                            oMat01.Columns.Item("CompltYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    //불량수량 검사
                    FailQty = 0;
                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        //불량코드를 입력했는지 check
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 && string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
                        {
                            errMessage = "불량수량이 입력되었을 때는 불량코드는 필수입니다.";
                            oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }

                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) == 0 && !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
                        {
                            errMessage = "불량코드를 확인하세요.";
                            oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }

                        if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value)
                         && (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value))
                        {
                            FailQty += Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value == "" ? "0" : oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value);
                        }
                    }

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) != FailQty)
                        {
                            errMessage = "공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.";
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111")
                    {
                        if (oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value == 1 && string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요";
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
                            errMessage = "비가동코드가 입력되었을 때는 비가동시간은 필수입니다.";
                            oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value) && oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value != "0")
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동시간이 입력되었을 때는 비가동코드는 필수입니다.";
                            oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (Convert.ToDouble(oMat02.Columns.Item("TTime").Cells.Item(i).Specific.Value == "" ? "0" : oMat02.Columns.Item("TTime").Cells.Item(i).Specific.Value) == 0)
                    {
                        oDS_PS_PP044M.SetValue("U_TTime", i - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value == "" ? "0" : oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value) 
                                                                                + Convert.ToDouble(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value == "" ? "0" : oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value)));
                    }
                }
                //비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E

                if (PS_PP044_Validate("검사01") == false)
                {
                    returnValue = false;
                    return returnValue;
                }

                oDS_PS_PP044L.RemoveRecord(oDS_PS_PP044L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP044M.RemoveRecord(oDS_PS_PP044M.Size - 1);
                oMat02.LoadFromDataSource();

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

            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP044_Validate(string ValidateType)
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

                if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value == "50"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value == "60") //작업타입이 일반,조정인경우
                {
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원인경우
                {
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30" || oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 외주, 실적인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();
                }
                
                if (ValidateType == "검사01")
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //입력된 행에 대해
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
                            Query01 = "  SELECT     PS_PP044H.DocEntry,";
                            Query01 += "            PS_PP044L.LineId,";
                            Query01 += "            CONVERT(NVARCHAR,PS_PP044H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP044L.LineId) AS DocInfo,";
                            Query01 += "            PS_PP044L.U_OrdGbn AS OrdGbn,";
                            Query01 += "            PS_PP044L.U_PP030HNo AS PP030HNo,";
                            Query01 += "            PS_PP044L.U_PP030MNo AS PP030MNo,";
                            Query01 += "            PS_PP044L.U_OrdMgNum AS OrdMgNum ";
                            Query01 += " FROM       [@PS_PP040H] PS_PP044H";
                            Query01 += "            LEFT JOIN";
                            Query01 += "            [@PS_PP040L] PS_PP044L";
                            Query01 += "                ON PS_PP044H.DocEntry = PS_PP044L.DocEntry ";
                            Query01 += " WHERE      PS_PP044L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                            RecordSet01.DoQuery(Query01);

                            for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                            {
                                Exist = false;
                                for (j = 1; j <= oMat01.VisualRowCount - 1; j++) //기존에 있는 행에대한 처리
                                {
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                    {
                                        //새로추가된 행인경우, 검사 불필요
                                    }
                                    else
                                    {
                                        if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value)
                                         && Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value)) //문서번호와 라인번호가 같으면 존재하는행
                                        {
                                            Exist = true;
                                        }
                                    }
                                }

                                if (Exist == false) //삭제된 행중 수량관계 확인
                                {
                                    if (RecordSet01.Fields.Item("OrdGbn").Value == "101") //휘팅
                                    {
                                        if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") //현재 공정 실적공정 여부
                                        {
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0 //휘팅벌크포장
                                             || Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) //휘팅실적
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
                                            //해당공정이 첫공정이면 입력 가능
                                        }
                                        else
                                        {
                                            PrevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP044H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'", 0, 1));
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
                                            CurrentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP044L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'", 0, 1));
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
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value))
                                {
                                    //새로추가된 행인경우, 검사 불필요
                                }
                                else
                                {
                                    if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "101") //휘팅
                                    {
                                        //현재 공정이 실적공정이면
                                        //현재 공정이 바렐 앞공정이면..
                                        if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oForm.Items.Item("DocEntry").Specific.Value + "-" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                                        {
                                            //휘팅벌크포장, 휘팅실적
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0
                                             || Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                                            {
                                                //작업일보등록된문서중에 수정이 된문서를 구함
                                                Query01 = "  SELECT     PS_PP044L.U_OrdMgNum,";
                                                Query01 += "            PS_PP044L.U_Sequence,";
                                                Query01 += "            PS_PP044L.U_CpCode,";
                                                Query01 += "            PS_PP044L.U_ItemCode,";
                                                Query01 += "            PS_PP044L.U_PP030HNo,";
                                                Query01 += "            PS_PP044L.U_PP030MNo,";
                                                Query01 += "            PS_PP044L.U_PQty,";
                                                Query01 += "            PS_PP044L.U_NQty,";
                                                Query01 += "            PS_PP044L.U_ScrapWt,";
                                                Query01 += "            PS_PP044L.U_WorkTime";
                                                Query01 += " FROM       [@PS_PP040H] PS_PP044H";
                                                Query01 += "            LEFT JOIN";
                                                Query01 += "            [@PS_PP040L] PS_PP044L";
                                                Query01 += "                ON PS_PP044H.DocEntry = PS_PP044L.DocEntry";
                                                Query01 += " WHERE      PS_PP044H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                                Query01 += "            AND PS_PP044L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                                Query01 += "            AND PS_PP044H.Canceled = 'N'";
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
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원인경우
                    {
                        //현재는 특별한 조건이 필요치 않음
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주인경우
                    {
                        //현재는 특별한 조건이 필요치 않음
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적인경우
                    {
                        //현재는 특별한 조건이 필요치 않음
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경우
                    {
                        //현재는 특별한 조건이 필요치 않음
                    }
                }
                else if (ValidateType == "행삭제01") //행삭제전 행삭제가능여부검사
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))
                        {
                            //새로추가된 행인경우, 삭제 가능
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") //휘팅
                            {
                                //현재 공정이 실적공정이면
                                //현재 공정이 바렐 앞공정이면..
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0 //휘팅벌크포장
                                     || Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0) //휘팅실적
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
                                Query01 += " GROUP BY   A.ItemCode";

                                string stockQty = string.IsNullOrEmpty(dataHelpClass.GetValue(Query01, 0, 1)) ? "0" : dataHelpClass.GetValue(Query01, 0, 1);

                                if (Convert.ToInt32(stockQty) > 0)
                                {
                                    errMessage = "재고가 존재하는 작번입니다. 삭제할 수 없습니다.";
                                    throw new Exception();
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
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") //작업타입이 조정인경
                    {
                    }
                }
                else if (ValidateType == "수정01") //수정전 수정가능여부검사
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") //작업타입이 일반인경우
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))
                        {
                            //새로추가된 행인경우, 수정 가능
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") ////휘팅
                            {
                                //현재 공정이 실적공정이면
                                //현재 공정이 바렐 앞공정이면..
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0 //휘팅벌크포장
                                     || Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0) //휘팅실적
                                    {
                                        errMessage = "수정된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
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
                        Query01 = "  SELECT     PS_PP044H.DocEntry,";
                        Query01 += "            PS_PP044L.LineId,";
                        Query01 += "            CONVERT(NVARCHAR,PS_PP044H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP044L.LineId) AS DocInfo,";
                        Query01 += "            PS_PP044L.U_OrdGbn AS OrdGbn,";
                        Query01 += "            PS_PP044L.U_PP030HNo AS PP030HNo,";
                        Query01 += "            PS_PP044L.U_PP030MNo AS PP030MNo,";
                        Query01 += "            PS_PP044L.U_OrdMgNum AS OrdMgNum ";
                        Query01 += " FROM       [@PS_PP040H] PS_PP044H";
                        Query01 += "            LEFT JOIN";
                        Query01 += "            [@PS_PP040L] PS_PP044L";
                        Query01 += "                ON PS_PP044H.DocEntry = PS_PP044L.DocEntry ";
                        Query01 += " WHERE      PS_PP044L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                        RecordSet01.DoQuery(Query01);

                        for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                        {
                            if (RecordSet01.Fields.Item("OrdGbn").Value == "101") //휘팅
                            {
                                //현재 공정이 실적이면
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y")
                                {
                                    if (Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0 //휘팅벌크포장
                                     || Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) //휘팅실적
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
                                //삭제된 행에 대한 검사..
                                OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
                                CurrentCpInfo = OrdMgNum;

                                PrevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'", 0, 1);
                                if (string.IsNullOrEmpty(PrevCpInfo))
                                {
                                    //해당공정이 첫공정이면 입력 가능
                                }
                                else
                                {
                                    PrevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP044H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevMATRIXCpQty = 0;

                                    CurrentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP044L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'", 0, 1));
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
            finally
            {

            }

            return returnValue;
        }

        /// <summary>
        /// 메트릭스에 데이터 로드(Mat04)
        /// </summary>
        private void PS_PP044_MTX01()
        {
            string OrdGbn;
            string BPLID;
            string DocDate;
            string Gubun;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);
            
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();

                sQry = "EXEC PS_PP044_01 '" + BPLID + "','" + DocDate + "', '" + OrdGbn + "', '" + Gubun + "'";
                oRecordSet01.DoQuery(sQry);

                oMat04.Clear();
                oDS_PS_PP044T.Clear();
                oMat04.FlushToDataSource();
                oMat04.LoadFromDataSource();

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP044T.Size)
                    {
                        oDS_PS_PP044T.InsertRecord(i);
                    }

                    oMat04.AddRow();
                    oDS_PS_PP044T.Offset = i;

                    oDS_PS_PP044T.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP044T.SetValue("U_ColReg01", i, "N");
                    oDS_PS_PP044T.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("CntcCode").Value.ToString().Trim());
                    oDS_PS_PP044T.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim());
                    oDS_PS_PP044T.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("Base").Value.ToString().Trim());
                    oDS_PS_PP044T.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("Extend").Value.ToString().Trim());
                    oDS_PS_PP044T.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("YTime").Value.ToString().Trim());
                    oDS_PS_PP044T.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("NTime").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                }

                oMat04.LoadFromDataSource();
                oMat04.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 메트릭스에 데이터 로드(Mat05)
        /// </summary>
        private void PS_PP044_MTX02()
        {
            string OrdGbn;
            string BPLID;
            string DocDate;
            string Gubun;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);

                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();

                sQry = "EXEC PS_PP044_02 '" + BPLID + "','" + DocDate + "', '" + OrdGbn + "', '" + Gubun + "'";
                oRecordSet01.DoQuery(sQry);

                oMat05.Clear();
                oDS_PS_PP044U.Clear();
                oMat05.FlushToDataSource();
                oMat05.LoadFromDataSource();

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP044U.Size)
                    {
                        oDS_PS_PP044U.InsertRecord(i);
                    }

                    oMat05.AddRow();
                    oDS_PS_PP044U.Offset = i;

                    oDS_PS_PP044U.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP044U.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("CntcCode").Value.ToString().Trim());
                    oDS_PS_PP044U.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim());
                    oDS_PS_PP044U.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("Base").Value.ToString().Trim());
                    oDS_PS_PP044U.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("Extend").Value.ToString().Trim());
                    oDS_PS_PP044U.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("YTime").Value.ToString().Trim());
                    oDS_PS_PP044U.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("NTime").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                }

                oMat05.LoadFromDataSource();
                oMat05.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 근무시간의 총합을 구함
        /// </summary>
        private void PS_PP044_SumWorkTime()
        {
            double total = 0;

            try
            {
                for (int loopCount = 0; loopCount <= oMat01.RowCount - 2; loopCount++)
                {
                    total += Convert.ToDouble(string.IsNullOrEmpty(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value) ? "0" : oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value);
                }

                oForm.Items.Item("Total").Specific.Value = total.ToString("#,##0.##");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
        }

        /// <summary>
        /// OrderInfoLoad
        /// </summary>
        private void PS_PP044_OrderInfoLoad()
        {
            string query;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value == "50"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value == "60"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value == "70") //일반,조정, 설계
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                    {
                        errMessage = "작업지시 관리번호를 입력하지 않습니다.";
                        throw new Exception();
                    }
                    else
                    {
                        query = "  SELECT   U_OrdGbn,";
                        query += "          U_BPLId,";
                        query += "          U_ItemCode,";
                        query += "          U_ItemName,";
                        query += "          U_OrdNum,";
                        query += "          U_OrdSub1,";
                        query += "          U_OrdSub2,";
                        query += "          DocEntry";
                        query += " FROM     [@PS_PP030H]";
                        query += " WHERE    U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'";
                        query += "          AND U_OrdGbn NOT IN ('104','107') ";
                        query += "          AND Canceled = 'N'";
                        RecordSet01.DoQuery(query);

                        if (RecordSet01.RecordCount == 0)
                        {
                            errMessage = "작업지시 정보가 존재하지 않습니다.";
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
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //PSMT
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                    {
                        errMessage = "작업지시 관리번호를 입력하지 않습니다.";
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
                        PS_PP044_AddMatrixRow01(0, true);

                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP044_AddMatrixRow02(0, true);

                        oMat03.Clear();
                        oMat03.FlushToDataSource();
                        oMat03.LoadFromDataSource();
                        oForm.Update();
                    }
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30")
                {
                    errMessage = "외주는 입력할 수 없습니다.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40")
                {
                    errMessage = "실적은 입력할 수 없습니다.";
                    throw new Exception();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// FindValidateDocument : 작업일보등록(방산부품) 문서인지 조회
        /// </summary>
        /// <param name="ObjectType"></param>
        /// <returns></returns>
        private bool PS_PP044_FindValidateDocument(string ObjectType)
        {
            bool returnValue = false;
            string query;
            string docEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                docEntry = oForm.Items.Item("DocEntry").Specific.Value;

                query = "  SELECT   DocEntry";
                query += " FROM     [" + ObjectType + "]";
                query += " WHERE    DocEntry = " + docEntry;

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
            string CntcCode;
            string IsYN;
            double YTime;

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP044_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oOrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                            oSequence = oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.Value.ToString().Trim();
                            oDocdate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                            oGubun = oForm.Items.Item("Gubun").Specific.Selected.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP044_DataValidCheck() == false)
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
                    else if (pVal.ItemUID == "2") //취소버튼 누를시 저장할 자료가 있으면 메시지 표시
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
                    else if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP044_OrderInfoLoad();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            PS_PP044_OrderInfoLoad();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }                    
                    else if (pVal.ItemUID == "Button03")
                    {
                        IsYN = "N";
                        if (oMat04.VisualRowCount == 0)
                        {
                            PSH_Globals.SBO_Application.MessageBox("작업자정보 라인이 존재하지 않습니다.");
                        }
                        else
                        {
                            for (int i = 1; i <= oMat04.VisualRowCount; i++)
                            {
                                if (oMat04.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                                {
                                    CntcCode = oMat04.Columns.Item("CntcCode").Cells.Item(i).Specific.Value;
                                    YTime = Convert.ToDouble(oMat04.Columns.Item("Base").Cells.Item(i).Specific.Value) + Convert.ToDouble(oMat04.Columns.Item("Extend").Cells.Item(i).Specific.Value);

                                    for (int j = 1; j <= oMat02.VisualRowCount - 1; j++)
                                    {
                                        if (CntcCode == oMat02.Columns.Item("WorkCode").Cells.Item(j).Specific.Value)
                                        {
                                            IsYN = "Y";
                                        }
                                    }
                                    if (IsYN == "N")
                                    {
                                        oDS_PS_PP044M.SetValue("U_TTime", oMat02.VisualRowCount - 1, Convert.ToString(YTime));
                                        oMat02.Columns.Item("YTime").Cells.Item(oMat02.VisualRowCount).Specific.Value = Convert.ToString(YTime);
                                        oMat02.Columns.Item("WorkCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = CntcCode;
                                    }
                                    IsYN = "N";
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "LBtn01") // 작업지시번호 링크 버튼
                    {
                        PS_PP030 tempForm = new PS_PP030();
                        tempForm.LoadForm(oForm.Items.Item("PP030HNo").Specific.Value);
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
                                    PS_PP044_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else
                                {
                                    PS_PP044_FormItemEnabled();
                                    PS_PP044_AddMatrixRow01(0, true);
                                    PS_PP044_AddMatrixRow02(0, true);
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
                                    PS_PP044_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP044_FormItemEnabled();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
                        if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10"
                         || oForm.Items.Item("OrdType").Specific.Selected.Value == "50"
                         || oForm.Items.Item("OrdType").Specific.Selected.Value == "60") //작업타입이 일반,조정일때
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "OrdMgNum", "");
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {

                            if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10"
                             || oForm.Items.Item("OrdType").Specific.Selected.Value == "50"
                             || oForm.Items.Item("OrdType").Specific.Selected.Value == "60"
                             || oForm.Items.Item("OrdType").Specific.Selected.Value == "70") //일반,조정, 설계
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
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
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //지원
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CItemCod");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UseMCode", "");
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //위쪽 방향키
                        {
                            if (pVal.Row > 1 & pVal.Row <= oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.CharPressed == 40) //아래 방향키
                        {
                            if (pVal.Row > 0 & pVal.Row < oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }

                        //작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
                        if (pVal.ColUID == "WorkTime" && pVal.Row != 0)
                        {
                            PS_PP044_SumWorkTime();
                        }
                    }
                    else if (pVal.ItemUID == "BaseTime")
                    {

                        //탭 키 Press
                        if (pVal.CharPressed == 9)
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
            string OrdGbn;
            int i;
            string sQry;
            int sCount;
            int sSeq;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                                //기타작업
                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP044_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                //기타작업
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP044M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP044_AddMatrixRow (pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                            }
                            else
                            {
                                oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "OrdType")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                
                                if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "10" 
                                 || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "50" 
                                 || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "60" 
                                 || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "70") //일반,조정,설계
                                {
                                    //창원은 품목구분 선택하도록 수정 '2015/04/09
                                    if (oForm.Items.Item("BPLId").Specific.Value == "1")
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
                                else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "20")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = true;
                                    oForm.Items.Item("BPLId").Enabled = true;
                                    oForm.Items.Item("ItemCode").Enabled = true;
                                }
                                else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "30")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("ItemCode").Enabled = false;
                                }
                                else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "40")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("ItemCode").Enabled = false;
                                }

                                oForm.Items.Item("OrdGbn").Specific.Select("102", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
                                PS_PP044_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP044_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);

                                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();

                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                oForm.Items.Item("OrdNum").Specific.Value = "";
                                oForm.Items.Item("ItemCode").Specific.Value = "";
                                oForm.Items.Item("PP030HNo").Specific.Value = "";
                                oForm.Items.Item("ItemName").Specific.Value = "";
                                oForm.Items.Item("OrdSub1").Specific.Value = "";
                                oForm.Items.Item("OrdSub2").Specific.Value = "";

                                sCount = oForm.Items.Item("Gubun").Specific.ValidValues.Count;
                                sSeq = sCount;
                                for (i = 1; i <= sCount; i++)
                                {
                                    oForm.Items.Item("Gubun").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                    sSeq -= 1;
                                }

                                sQry = "SELECT U_Minor, U_CdName From [@PS_SY001L] Where Code = 'P208' and U_RelCd = '" + OrdGbn + "' Order by U_Minor";
                                oRecordSet01.DoQuery(sQry);
                                oForm.Items.Item("Gubun").Specific.ValidValues.Add("선택", "선택");
                                while (!oRecordSet01.EoF)
                                {
                                    oForm.Items.Item("Gubun").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet01.MoveNext();
                                }
                                oForm.Items.Item("Gubun").Specific.Select("선택", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP044_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP044_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                oMat04.LoadFromDataSource();

                                oMat05.Clear();
                                oMat05.FlushToDataSource();
                                oMat05.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "Gubun")
                            {
                                PS_PP044_MTX01();
                                PS_PP044_MTX02();
                            }
                            else
                            {
                                if (pVal.ItemUID != "Gubun")
                                {
                                    oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                }
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oMat02.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat03.LoadFromDataSource();
                        oMat03.AutoResizeColumns();
                        oMat04.LoadFromDataSource();
                        oMat04.AutoResizeColumns();
                        oMat05.LoadFromDataSource();
                        oMat05.AutoResizeColumns();
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
                        else if (pVal.ItemUID == "Mat04")
                        {
                            oMat04.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
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
                    else if (pVal.ItemUID == "Opt03")
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
                            if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value == "50" || oForm.Items.Item("OrdType").Specific.Selected.Value == "60") //작업타입이 일반,조정인경우
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                }
                                else
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP044_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP044_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }

                                    oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
                            }
                            else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원인경우
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                }
                                else
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP044_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP044_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }

                                    oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
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
                        if (pVal.ColUID == "PP030HNo")
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
            double Weight;
            double Time;
            double hour;
            double minute;
            string WkCmDt;
            string errCode = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = null;
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
                            if (PS_PP044_Validate("수정01") == false)
                            {
                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrdMgNum")
                                {
                                    RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) //작지번호에 값이 없으면 작업지시가 불러오기전
                                    {
                                        oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                    }
                                    else //작업지시가 선택된상태
                                    {
                                        if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10"
                                         || oForm.Items.Item("OrdType").Specific.Selected.Value == "50"
                                         || oForm.Items.Item("OrdType").Specific.Selected.Value == "60"
                                         || oForm.Items.Item("OrdType").Specific.Selected.Value == "70") //작업타입이 일반,조정, 설계
                                        {
                                            //작지문서헤더번호가 일치하지 않으면
                                            if (oForm.Items.Item("PP030HNo").Specific.Value != codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")))
                                            {
                                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else //작지문서번호가 일치하면
                                            {
                                                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1") //신동사업부를 제외한 사업부만 체크
                                                {   
                                                    for (i = 1; i <= oMat01.RowCount; i++)
                                                    {
                                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                                        {
                                                            PSH_Globals.SBO_Application.MessageBox("이미 입력한 공정입니다.");
                                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                            errCode = "1";
                                                            throw new Exception();
                                                        }
                                                    }

                                                    //생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_S
                                                    query01 = "EXEC PS_PP040_90 '";
                                                    query01 += codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")) + "'";
                                                    RecordSet01.DoQuery(query01);
                                                    WkCmDt = RecordSet01.Fields.Item("WkCmDt").Value;

                                                    //생산완료수량이 작업지시수량만큼 모두 등록이 되었다면
                                                    if (RecordSet01.Fields.Item("Return").Value == "1")
                                                    {
                                                        if (PSH_Globals.SBO_Application.MessageBox("생산완료가 모두 등록된 작번(완료일자:" + WkCmDt + ")입니다. 계속 진행하시겠습니까?", 1, "예", "아니오") == 1)
                                                        {
                                                            //계속 진행시에는 해당 작업지시문서번호 등록
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                            errCode = "1";
                                                            throw new Exception();
                                                        }
                                                    }
                                                    //생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_E
                                                }

                                                query01 = "EXEC PS_PP040_01 '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "', '" + oForm.Items.Item("OrdType").Specific.Selected.Value + "'";
                                                RecordSet01.DoQuery(query01);
                                                if (RecordSet01.RecordCount == 0)
                                                {
                                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                }
                                                else
                                                {
                                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                    oDS_PS_PP044L.SetValue("U_Sequence", pVal.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
                                                    oDS_PS_PP044L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                    oDS_PS_PP044L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
                                                    oDS_PS_PP044L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
                                                    oDS_PS_PP044L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
                                                    oDS_PS_PP044L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                    oDS_PS_PP044L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                                    oDS_PS_PP044L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
                                                    oDS_PS_PP044L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
                                                    oDS_PS_PP044L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
                                                    oDS_PS_PP044L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
                                                    oDS_PS_PP044L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
                                                    oDS_PS_PP044L.SetValue("U_SelWt", pVal.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
                                                    oDS_PS_PP044L.SetValue("U_PSum", pVal.Row - 1, RecordSet01.Fields.Item("PSum").Value);
                                                    oDS_PS_PP044L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                    oDS_PS_PP044L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                    oDS_PS_PP044L.SetValue("U_LineId", pVal.Row - 1, "");

                                                    //설비코드,명 Reset
                                                    oDS_PS_PP044L.SetValue("U_MachCode", pVal.Row - 1, "");
                                                    oDS_PS_PP044L.SetValue("U_MachName", pVal.Row - 1, "");

                                                    if (oMat03.VisualRowCount == 0) //불량코드테이블
                                                    {
                                                        PS_PP044_AddMatrixRow03(0, true);
                                                    }
                                                    else
                                                    {
                                                        PS_PP044_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                    }

                                                    oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                    oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                    oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
                                                    oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));

                                                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50" || oForm.Items.Item("OrdType").Specific.Selected.Value == "60")
                                                    {
                                                        oDS_PS_PP044H.SetValue("U_BaseTime", 0, "1");
                                                        oMat02.Columns.Item("WorkCode").Cells.Item(1).Specific.Value = "9999999";
                                                        oDS_PS_PP044M.SetValue("U_WorkName", 0, "조정");
                                                        oMat02.LoadFromDataSource();
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                            }
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") //작업타입이 PSMT지원
                                        {
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) == 0) //올바른 공정코드인지 검사
                                            {
                                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else
                                            {
                                                for (i = 1; i <= oMat01.RowCount; i++)
                                                {

                                                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                                    {
                                                        PSH_Globals.SBO_Application.StatusBar.SetText("이미 입력한 공정입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                                        oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                        errCode = "1";
                                                        throw new Exception();
                                                    }
                                                }
                                                oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_CpCode", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_CpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                                oDS_PS_PP044L.SetValue("U_OrdGbn", pVal.Row - 1, oForm.Items.Item("OrdGbn").Specific.Selected.Value);
                                                oDS_PS_PP044L.SetValue("U_BPLId", pVal.Row - 1, oForm.Items.Item("BPLId").Specific.Selected.Value);
                                                oDS_PS_PP044L.SetValue("U_ItemCode", pVal.Row - 1, "");
                                                oDS_PS_PP044L.SetValue("U_ItemName", pVal.Row - 1, "");
                                                oDS_PS_PP044L.SetValue("U_OrdNum", pVal.Row - 1, oForm.Items.Item("OrdNum").Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_OrdSub1", pVal.Row - 1, oForm.Items.Item("OrdSub1").Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_OrdSub2", pVal.Row - 1, oForm.Items.Item("OrdSub2").Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_PP030HNo", pVal.Row - 1, "");
                                                oDS_PS_PP044L.SetValue("U_PP030MNo", pVal.Row - 1, "");
                                                oDS_PS_PP044L.SetValue("U_PSum", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP044L.SetValue("U_ScrapWt", pVal.Row - 1, "0");

                                                if (oMat03.VisualRowCount == 0) //불량코드테이블
                                                {
                                                    PS_PP044_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    if (oDS_PS_PP044L.GetValue("U_OrdMgNum", pVal.Row - 1) == oDS_PS_PP044N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1))
                                                    {
                                                    }
                                                    else
                                                    {
                                                        PS_PP044_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                    }
                                                }
                                                oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                            }
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") //작업타입이 외주
                                        {
                                        }
                                        else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") //작업타입이 실적
                                        {

                                        }

                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP044_AddMatrixRow01(pVal.Row, false);
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "PQty")
                                {
                                    Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;

                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                            if (Weight == 0)
                                            {
                                                oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            }
                                            else
                                            {
                                                oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, "0");
                                            oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                        }
                                        else
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else
                                    {
                                        oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                        if (Weight == 0)
                                        {
                                            oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else
                                        {
                                            oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }
                                else if (pVal.ColUID == "NQty")
                                {
                                    Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;

                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));

                                            if (Weight == 0)
                                            {
                                                oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            else
                                            {
                                                oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                            }
                                        }
                                        else
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));

                                            if (Weight == 0)
                                            {
                                                oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            else
                                            {
                                                oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                            }
                                        }
                                        else
                                        {
                                            oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else
                                    {
                                        oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));

                                        if (Weight == 0)
                                        {
                                            oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else
                                        {
                                            oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "WorkTime")
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                                else if (pVal.ColUID == "BdwQty") //기준도면매수
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100));
                                    oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "DwRate") //도면 적용률
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100));
                                    oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "NdwQTy") //신규도면매수
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "MachCode")
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "CItemCod")
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP044L.SetValue("U_CItemNam", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "WorkCode")
                            {
                                //기타작업
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP044M.SetValue("U_WorkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP044M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP044_AddMatrixRow02(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "NStart")
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        Time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        Time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    hour = Time / 100;
                                    minute = Time % 100;
                                    Time = hour;
                                    if (minute > 0)
                                    {
                                        Time += +0.5;
                                    }
                                    oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
                                    oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                    oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                }
                            }
                            else if (pVal.ColUID == "NEnd")
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        Time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        Time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    hour = Time / 100;
                                    minute = Time % 100;
                                    Time = hour;
                                    if (minute > 0)
                                    {
                                        Time += 0.5;
                                    }
                                    oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
                                    oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                    oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                }
                            }
                            else if (pVal.ColUID == "YTime")
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                double tTime = Convert.ToDouble(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value == "" ? "0" : oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value);

                                if (tTime > 0)
                                {
                                    oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(tTime - Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                }
                            }
                            else if (pVal.ColUID == "NTime")
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                double tTime = Convert.ToDouble(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value == "" ? "0" : oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value);

                                if (tTime > 0)
                                {
                                    oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(tTime - Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                }
                            }
                            else
                            {
                                oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }

                            oMat02.LoadFromDataSource();
                            oMat02.AutoResizeColumns();
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP044N.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }

                            oMat03.LoadFromDataSource();
                            oMat03.AutoResizeColumns();
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP044H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BaseTime")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "OrdMgNum")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    PS_PP044_OrderInfoLoad();
                                }
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);

                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP044_AddMatrixRow01(0, true);

                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP044_AddMatrixRow02(0, true);

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
                            else if (pVal.ItemUID == "DocDate")
                            {
                                PS_PP044_MTX01();
                                PS_PP044_MTX02();
                            }
                            else
                            {
                                oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }

                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        oForm.Update();
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
                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                    {
                        PS_PP044_AddMatrixRow01(pVal.Row, false);
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
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (RecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                }

                oForm.Freeze(false);
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
                PS_PP044_SumWorkTime();

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP044_FormItemEnabled();
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_PP044_AddMatrixRow01(oMat01.VisualRowCount, false);
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        PS_PP044_AddMatrixRow02(oMat02.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat04);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat05);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044N);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044T);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP044U);
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
                    PS_PP044_FormResize();
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
                        PS_PP044_AddMatrixRow01(0, true);
                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP044_AddMatrixRow02(0, true);
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
                        if (oLastItemUID01 == "Mat01")
                        {
                            if (PS_PP044_Validate("행삭제01") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            for (i = 1; i <= oMat03.RowCount; i++)
                            {
                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                 && oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value)
                                {
                                    oDS_PS_PP044N.RemoveRecord(i - 1);
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
                                if (oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value != 1)
                                {
                                    oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value = oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value - 1;
                                }
                            }

                            oMat01.FlushToDataSource();
                            oDS_PS_PP044L.RemoveRecord(oDS_PS_PP044L.Size - 1);
                            oMat01.LoadFromDataSource();

                            if (oMat01.RowCount == 0)
                            {
                                PS_PP044_AddMatrixRow01(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP044L.GetValue("U_OrdMgNum", oMat01.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP044_AddMatrixRow01(oMat01.RowCount, false);
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
                            oDS_PS_PP044M.RemoveRecord(oDS_PS_PP044M.Size - 1);
                            oMat02.LoadFromDataSource();

                            if (oMat02.RowCount == 0)
                            {
                                PS_PP044_AddMatrixRow02(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP044M.GetValue("U_WorkCode", oMat02.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP044_AddMatrixRow02(oMat02.RowCount, false);
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

                            if (oDS_PS_PP044N.Size == 1)
                            {
                            }
                            else
                            {
                                oDS_PS_PP044N.RemoveRecord(oDS_PS_PP044N.Size - 1);
                            }
                            oMat03.LoadFromDataSource();

                            //공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                            for (i = 1; i <= oMat01.RowCount - 1; i++)
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
                                        PS_PP044_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP044_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }

                                    oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(i));
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
                        query01 += "                (SELECT MIN(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('102','602'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '10'";
                        query01 += "            AND U_OrdGbn IN ('102','602')";
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
                        query01 += "                (SELECT MAX(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('102','602'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '10'";
                        query01 += "            AND U_OrdGbn IN ('102','602')";
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
                    query01 += "            AND U_OrdGbn IN ('102','602')";

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
                    query01 += "            AND U_OrdGbn IN ('102','602')";

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
                                if (PS_PP044_Validate("취소") == false)
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
                            PS_PP044_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP044_FormItemEnabled();
                            PS_PP044_AddMatrixRow01(0, true);
                            PS_PP044_AddMatrixRow02(0, true);
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
                        if (PS_PP044_FindValidateDocument("@PS_PP040H") == false)
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
