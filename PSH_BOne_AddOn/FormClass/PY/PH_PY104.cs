using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 고정수당공제금액일괄등록
    /// </summary>
    internal class PH_PY104 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.Grid oGrid2;
        private SAPbouiCOM.DataTable oDS_PH_PY104_01;
        private SAPbouiCOM.DataTable oDS_PH_PY104_02;
        private int tSeqAll; //그리드1의 체크 순번
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY104.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY104_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY104");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY104_CreateItems();
                PH_PY104_EnableMenus();
                PH_PY104_SetDocument(oFormDocEntry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid1").Specific;
                oForm.DataSources.DataTables.Add("PH_PY104_01");
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("선택", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("순서", SAPbouiCOM.BoFieldsType.ft_Float);
                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_01");
                oDS_PH_PY104_01 = oForm.DataSources.DataTables.Item("PH_PY104_01");

                oGrid2 = oForm.Items.Item("Grid2").Specific;
                oForm.DataSources.DataTables.Add("PH_PY104_02");
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("체크", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_02");
                oDS_PH_PY104_02 = oForm.DataSources.DataTables.Item("PH_PY104_02");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 급여형태
                oForm.DataSources.UserDataSources.Add("PAYTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PAYTYP").Specific.DataBind.SetBound(true, "", "PAYTYP");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN= 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYTYP").Specific, "");
                oForm.Items.Item("PAYTYP").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PAYTYP").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("PAYTYP").DisplayDesc = true;

                // 직급형태From
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JIGCODF").Specific,"");
                oForm.Items.Item("JIGCODF").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("JIGCODF").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JIGCODF").DisplayDesc = true;

                // 직급형태To
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JIGCODT").Specific,"");
                oForm.Items.Item("JIGCODT").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("JIGCODT").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JIGCODT").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_EnableMenus
        /// </summary>
        private void PH_PY104_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                // 제거
                oForm.EnableMenu("1284", false);               // 취소
                oForm.EnableMenu("1293", true);                // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY104_SetDocument
        /// </summary>
        private void PH_PY104_SetDocument(string oFormDocEntry)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry)))
                {
                    PH_PY104_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY104_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY104_FormItemEnabled()
        {
            int i;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    // 기본사항 - 부서 (사업장에 따른 부서변경)
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx))
                    {
                        sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                        sQry += " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += " ORDER BY U_Code";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific,"");
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                        oForm.Items.Item("TeamCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    // 담당 (사업장에 따른 담당변경)
                    if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx))
                    {
                        sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                        sQry += " WHERE Code = '2' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += " Order By U_Code";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific,"");
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                        oForm.Items.Item("RspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    tSeqAll = 0;
                    PH_PY104_DataLoad();

                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", false); // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    // 부서
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "-");
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    // 담당
                    if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "-");
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    tSeqAll = 0;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY104_DataLoad();

                    oForm.EnableMenu("1281", false); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY104_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            int j = 0;

            try
            {
                if (oGrid1.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                    {
                        if (oGrid1.DataTable.GetValue("SLT", i) == "Y")
                        {
                            if (!string.IsNullOrEmpty(oGrid1.DataTable.GetValue("U_CSUCOD", i)))
                            {
                                j += 1;
                            }
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("수당, 공제 테이블에 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }

                if (j == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("선택된 수당, 공제 데이터가 없습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_DataFind()
        {
            int i;
            int iRow;
            string sQry;
            string CLTCOD;
            string TeamCode;
            string RspCode;
            string PAYTYP;
            string JIGCODF;
            string JIGCODT;
            string HOBONGF;
            string HOBONGT;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                // PH_PY104_TEMP 테이블 초기화
                sQry = "DELETE PH_PY104_TEMP";
                oRecordSet.DoQuery(sQry);
                // PH_PY104_TEMP2 테이블 초기화
                sQry = "DELETE PH_PY104_TEMP2";
                oRecordSet.DoQuery(sQry);

                // 그리드1 체크 데이터 PH_PY104_TEMP  저장
                if (oGrid1.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_01.GetValue("SLT", i) == "Y")
                        {
                            sQry = "EXEC PH_PY104_Grid1 '";
                            sQry += oDS_PH_PY104_01.GetValue("GBN", i) + "','";
                            sQry += oDS_PH_PY104_01.GetValue("U_CSUCOD", i) + "','";
                            sQry += oDS_PH_PY104_01.GetValue("U_CSUNAM", i) + "','";
                            sQry += oDS_PH_PY104_01.GetValue("SEQ", i) + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                    }
                }

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                PAYTYP = oForm.Items.Item("PAYTYP").Specific.Value.ToString().Trim();
                JIGCODF = (oForm.Items.Item("JIGCODF").Specific.Value.Trim() == "%" ? "00" : oForm.Items.Item("JIGCODF").Specific.Value.Trim());
                JIGCODT = (oForm.Items.Item("JIGCODT").Specific.Value.Trim() == "%" ? "ZZ" : oForm.Items.Item("JIGCODT").Specific.Value.Trim());
                HOBONGF = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGF").Specific.Value.Trim()) ? "000" : oForm.Items.Item("HOBONGF").Specific.Value.Trim());
                HOBONGT = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGT").Specific.Value.Trim()) ? "ZZZ" : oForm.Items.Item("HOBONGT").Specific.Value.Trim());

                // 검색 조건 - 임시 테이블 저장 PH_PY104_TEMP2
                sQry = "Exec PH_PY104_Grid2 '";
                sQry += CLTCOD + "','";
                sQry += TeamCode + "','";
                sQry += RspCode + "','";
                sQry += PAYTYP + "','";
                sQry += JIGCODF + "','";
                sQry += JIGCODT + "','";
                sQry += HOBONGF + "','";
                sQry += HOBONGT + "'";
                oRecordSet.DoQuery(sQry);

                // 그리드1 체크 데이터P PH_PY104_TEMP 불러옴
                sQry = "SELECT GUBUN, CSUCOD, CSUNAM FROM PH_PY104_TEMP ORDER BY SEQ";
                oRecordSet.DoQuery(sQry);

                // PH_PY104_TEMP2  데이터에 코드와 이름 붙여서 다시 셀렉트
                if (oRecordSet.RecordCount > 0)
                {
                    sQry = "SELECT '' AS ChkBox , T0.CODE, T0.Name ";
                    for (i = 1; i <= oRecordSet.RecordCount; i++)
                    {
                        if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "수당")
                        {
                            sQry += ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001B] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + oRecordSet.Fields.Item(2).Value + "'";

                        }
                        else if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "공제")
                        {
                            sQry += ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001C] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + oRecordSet.Fields.Item(2).Value + "'";
                        }
                        oRecordSet.MoveNext();
                    }
                    sQry += " FROM  PH_PY104_TEMP2 T0";
                    oDS_PH_PY104_02.ExecuteQuery(sQry);
                }
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY104_TitleSetting_Grid2(iRow);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_DataLoad
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_DataLoad()
        {
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                sQry = "  SELECT '수당' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
                sQry += " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code";
                sQry += " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY102A])";
                sQry += " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value + "'";
                sQry += " Union All";
                sQry += " SELECT '공제' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
                sQry += " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code";
                sQry += " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY103A])";
                sQry += " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value + "'";

                oDS_PH_PY104_01.ExecuteQuery(sQry);
                PH_PY104_TitleSetting_Grid1();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataLoad_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_DataCopy
        /// </summary>
        /// <returns></returns>
        private bool PH_PY104_DataCopy()
        {
            bool functionReturnValue = false;
            int i;
            int j;
            int First;
            object[] FirstData = null;
            int TOTCNT;
            int V_StatusCnt;
            int oProValue;
            int tRow;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                First = 0;
                if (oGrid2.Rows.Count > 0)
                {
                    ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                    // 최대값 구하기
                    TOTCNT = oGrid2.Rows.Count;
                    V_StatusCnt = TOTCNT / 50;
                    oProValue = 1;
                    tRow = 1;

                    for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y")
                        {
                            First += 1;
                            if (First == 1)
                            {
                                FirstData = new object[oGrid2.Columns.Count + 1];
                                for (j = 0; j <= oGrid2.Columns.Count - 1; j++)
                                {
                                    FirstData[j] = oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value;
                                }
                            }
                            else
                            {
                                for (j = 3; j <= oGrid2.Columns.Count - 1; j++)
                                {
                                    oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value = FirstData[j];
                                }
                            }
                        }

                        tRow += 1;
                        if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                        {
                            ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                            oProValue += 1;
                            ProgressBar01.Value = oProValue;
                        }
                    }
                }

                if (First == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("선택된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("선택된 필드의 전체 값 복사가 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    functionReturnValue = true;
                }
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataCopy_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }

                oForm.Freeze(false);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// DataSave
        /// </summary>
        /// <returns></returns>
        private bool PH_PY104_DataSave()
        {
            bool functionReturnValue = false;
            int i;
            int j;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oGrid2.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y")
                        {
                            for (j = 3; j <= oGrid2.Columns.Count - 1; j++)
                            {
                                sQry = "SELECT GUBUN FROM PH_PY104_TEMP WHERE CSUNAM = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                oRecordSet.DoQuery(sQry);

                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "수당")
                                    {
                                        sQry = "UPDATE [@PH_PY001B] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
                                        sQry += " WHERE Code = '" + oDS_PH_PY104_02.GetValue("CODE", i).ToString().Trim() + "'";
                                        sQry += " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                        oRecordSet.DoQuery(sQry);
                                    }
                                    else if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "공제")
                                    {
                                        sQry = "UPDATE [@PH_PY001C] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
                                        sQry += " WHERE Code = '" + oDS_PH_PY104_02.GetValue("CODE", i).ToString().Trim() + "'";
                                        sQry += " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                        oRecordSet.DoQuery(sQry);
                                    }
                                }
                            }
                            PSH_Globals.SBO_Application.SetStatusBarMessage("[" + oDS_PH_PY104_02.GetValue("CODE", i).ToString().Trim() + "] 의 수당,공제 데이터가 갱신중입니다..", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    functionReturnValue = true;
                    PSH_Globals.SBO_Application.SetStatusBarMessage("수당,공제 데이터가 갱신 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataSave_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY104_TitleSetting_Grid1 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_TitleSetting_Grid1()
        {
            int i;
            string[] COLNAM = new string[5];

            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "구분";
                COLNAM[1] = "코드";
                COLNAM[2] = "코드명";
                COLNAM[3] = "선택";
                COLNAM[4] = "순서";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i <= 2)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    else if (i == 3)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    }
                    else if (i == 4)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid1_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_TitleSetting_Grid2 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_TitleSetting_Grid2(int iRoW)
        {
            int i;
            string[] COLNAM = new string[3];

            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "체크";
                COLNAM[1] = "사번";
                COLNAM[2] = "이름";

                for (i = 0; i <= oGrid2.Columns.Count - 1; i++)
                {
                    if (i == 0)
                    {
                        oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                        oGrid2.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                        oGrid2.Columns.Item(i).Editable = true;
                    }
                    else if (i == 1 | i == 2)
                    {
                        oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                        oGrid2.Columns.Item(i).Editable = false;
                    }
                    else
                    {
                        oGrid2.Columns.Item(i).Editable = true;
                        oGrid2.Columns.Item(i).RightJustified = true;
                    }
                }
                oGrid2.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid2_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Check_Seq 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void Check_Seq(string ColUID, int Row)
        {
            int i;
            int tSeq;

            try
            {
                oForm.Freeze(true);

                if (oGrid1.Rows.Count > 0 & Row >= 0)
                {
                    if (tSeqAll < 0)
                    {
                        tSeqAll = 0;
                    }

                    if (oGrid1.DataTable.GetValue("SLT", Row) == "Y")
                    {
                        tSeqAll += 1;
                        oGrid1.DataTable.SetValue("SEQ", Row, tSeqAll);
                    }
                    else if (oGrid1.DataTable.GetValue("SLT", Row) == "N")
                    {
                        tSeqAll -= 1;
                        if (string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", Row)))
                        {
                            oGrid1.DataTable.SetValue("SEQ", Row, 0);
                        }

                        tSeq = oGrid1.DataTable.GetValue("SEQ", Row);
                        oGrid1.DataTable.SetValue("SEQ", Row, "");

                        for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                        {
                            if (oGrid1.DataTable.GetValue("SEQ", i) > tSeq & !string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", i)))
                            {
                                oGrid1.DataTable.SetValue("SEQ", i, Convert.ToInt16(oGrid1.DataTable.GetValue("SEQ", i)) - 1);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Check_Seq_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS://                    4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                    ////7
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                    ////8
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:                    ////9
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                    ////12
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                    ////16
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    ////18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    ////19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                    ////20
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    ////22
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                    ////23
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                    ////37
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                    ////38
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag:                    ////39
                //    break;

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
            int i;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        if (PH_PY104_DataValidCheck() == true)
                        {
                            PH_PY104_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        if (PH_PY104_DataSave() == false)
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_AllChk")
                    {
                        if (oGrid2.Rows.Count > 0)
                        {
                            oForm.Freeze(true);
                            for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                            {
                                oDS_PH_PY104_02.SetValue("ChkBox", i, "Y");
                            }
                            oForm.Freeze(false);
                        }
                    }
                    if (pVal.ItemUID == "Btn_Copy")
                    {
                        PH_PY104_DataCopy();
                    }

                }
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_GOT_FOCUS
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.ItemUID)
                {
                    case "Grid1":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_COMBO_SELECT
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {

                        if (pVal.ItemUID == "CLTCOD")
                        {
                            // 기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += " ORDER BY U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific,"");
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                            oForm.Items.Item("TeamCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("TeamCode").DisplayDesc = true;

                            PH_PY104_DataLoad();

                            // 담당 (사업장에 따른 담당변경)
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }
                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '2' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += " Order By U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific,"");
                            oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                            oForm.Items.Item("RspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("RspCode").DisplayDesc = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid1":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ActionSuccess == true)
                    {
                        if (pVal.ItemUID == "Grid1" & pVal.ColUID == "SLT")
                        {
                            Check_Seq(pVal.ColUID, pVal.Row);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY104_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY104_01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY104_02);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY104_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY104_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY104_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY104_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
