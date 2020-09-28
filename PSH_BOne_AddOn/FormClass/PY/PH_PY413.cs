using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 월세액.주택임차차입금자료 등록
    /// </summary>
    internal class PH_PY413 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //  그리드 사용시
        // public SAPbouiCOM.Grid oGrid1;
        // public SAPbouiCOM.DataTable oDS_PH_PY413;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY413.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY413_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY413");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY413_CreateItems();
                PH_PY413_FormItemEnabled();
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY413_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY413");

                // oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY413");
                // oDS_PH_PY413 = oForm.DataSources.DataTables.Item("PH_PY413");

                // 그리드 타이틀 
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("순번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("구분코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("구분명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("금융기관코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("금융기관명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("계좌번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("납입년차", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("납입금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("공제금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("투자년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PH_PY413").Columns.Add("투자구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 임대인성명
                oForm.DataSources.UserDataSources.Add("ws_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ws_name1").Specific.DataBind.SetBound(true, "", "ws_name1");

                oForm.DataSources.UserDataSources.Add("ws_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ws_name2").Specific.DataBind.SetBound(true, "", "ws_name2");

                oForm.DataSources.UserDataSources.Add("ws_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ws_name3").Specific.DataBind.SetBound(true, "", "ws_name3");

                // 임대인주민등록번호
                oForm.DataSources.UserDataSources.Add("ws_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ws_jumin1").Specific.DataBind.SetBound(true, "", "ws_jumin1");

                oForm.DataSources.UserDataSources.Add("ws_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ws_jumin2").Specific.DataBind.SetBound(true, "", "ws_jumin2");

                oForm.DataSources.UserDataSources.Add("ws_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ws_jumin3").Specific.DataBind.SetBound(true, "", "ws_jumin3");

                // 주택유형
                oForm.DataSources.UserDataSources.Add("ws_hcode1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ws_hcode1").Specific.DataBind.SetBound(true, "", "ws_hcode1");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ws_hcode1").Specific, "Y");
                oForm.Items.Item("ws_hcode1").DisplayDesc = true;
                oForm.Items.Item("ws_hcode1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Add("ws_hcode2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ws_hcode2").Specific.DataBind.SetBound(true, "", "ws_hcode2");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ws_hcode2").Specific, "Y");
                oForm.Items.Item("ws_hcode2").DisplayDesc = true;
                oForm.Items.Item("ws_hcode2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Add("ws_hcode3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ws_hcode3").Specific.DataBind.SetBound(true, "", "ws_hcode3");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ws_hcode3").Specific, "Y");
                oForm.Items.Item("ws_hcode3").DisplayDesc = true;
                oForm.Items.Item("ws_hcode3").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 주택계약면적
                oForm.DataSources.UserDataSources.Add("ws_hm1", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ws_hm1").Specific.DataBind.SetBound(true, "", "ws_hm1");

                oForm.DataSources.UserDataSources.Add("ws_hm2", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ws_hm2").Specific.DataBind.SetBound(true, "", "ws_hm2");

                oForm.DataSources.UserDataSources.Add("ws_hm3", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ws_hm3").Specific.DataBind.SetBound(true, "", "ws_hm3");

                // 임대차계약서상주소
                oForm.DataSources.UserDataSources.Add("ws_addr1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ws_addr1").Specific.DataBind.SetBound(true, "", "ws_addr1");

                oForm.DataSources.UserDataSources.Add("ws_addr2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ws_addr2").Specific.DataBind.SetBound(true, "", "ws_addr2");

                oForm.DataSources.UserDataSources.Add("ws_addr3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ws_addr3").Specific.DataBind.SetBound(true, "", "ws_addr3");

                // 임대차계약기간
                oForm.DataSources.UserDataSources.Add("ws_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_fymd1").Specific.DataBind.SetBound(true, "", "ws_fymd1");

                oForm.DataSources.UserDataSources.Add("ws_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_fymd2").Specific.DataBind.SetBound(true, "", "ws_fymd2");

                oForm.DataSources.UserDataSources.Add("ws_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_fymd3").Specific.DataBind.SetBound(true, "", "ws_fymd3");

                oForm.DataSources.UserDataSources.Add("ws_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_tymd1").Specific.DataBind.SetBound(true, "", "ws_tymd1");

                oForm.DataSources.UserDataSources.Add("ws_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_tymd2").Specific.DataBind.SetBound(true, "", "ws_tymd2");

                oForm.DataSources.UserDataSources.Add("ws_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_tymd3").Specific.DataBind.SetBound(true, "", "ws_tymd3");

                // 월세액
                oForm.DataSources.UserDataSources.Add("ws_mamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_mamt1").Specific.DataBind.SetBound(true, "", "ws_mamt1");

                oForm.DataSources.UserDataSources.Add("ws_mamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_mamt2").Specific.DataBind.SetBound(true, "", "ws_mamt2");

                oForm.DataSources.UserDataSources.Add("ws_mamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_mamt3").Specific.DataBind.SetBound(true, "", "ws_mamt3");

                // 공제금액
                oForm.DataSources.UserDataSources.Add("ws_gamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_gamt1").Specific.DataBind.SetBound(true, "", "ws_gamt1");

                oForm.DataSources.UserDataSources.Add("ws_gamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_gamt2").Specific.DataBind.SetBound(true, "", "ws_gamt2");

                oForm.DataSources.UserDataSources.Add("ws_gamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_gamt3").Specific.DataBind.SetBound(true, "", "ws_gamt3");

                // 금전소비대차 계약내용
                // 대주
                oForm.DataSources.UserDataSources.Add("dj_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("dj_name1").Specific.DataBind.SetBound(true, "", "dj_name1");

                oForm.DataSources.UserDataSources.Add("dj_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("dj_name2").Specific.DataBind.SetBound(true, "", "dj_name2");

                oForm.DataSources.UserDataSources.Add("dj_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("dj_name3").Specific.DataBind.SetBound(true, "", "dj_name3");

                // 대주 주민등록번호
                oForm.DataSources.UserDataSources.Add("dj_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("dj_jumin1").Specific.DataBind.SetBound(true, "", "dj_jumin1");

                oForm.DataSources.UserDataSources.Add("dj_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("dj_jumin2").Specific.DataBind.SetBound(true, "", "dj_jumin2");

                oForm.DataSources.UserDataSources.Add("dj_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("dj_jumin3").Specific.DataBind.SetBound(true, "", "dj_jumin3");

                // 금전소비대차계약기간
                oForm.DataSources.UserDataSources.Add("dj_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_fymd1").Specific.DataBind.SetBound(true, "", "dj_fymd1");

                oForm.DataSources.UserDataSources.Add("dj_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_fymd2").Specific.DataBind.SetBound(true, "", "dj_fymd2");

                oForm.DataSources.UserDataSources.Add("dj_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_fymd3").Specific.DataBind.SetBound(true, "", "dj_fymd3");

                oForm.DataSources.UserDataSources.Add("dj_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_tymd1").Specific.DataBind.SetBound(true, "", "dj_tymd1");

                oForm.DataSources.UserDataSources.Add("dj_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_tymd2").Specific.DataBind.SetBound(true, "", "dj_tymd2");

                oForm.DataSources.UserDataSources.Add("dj_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_tymd3").Specific.DataBind.SetBound(true, "", "dj_tymd3");

                // 차임급이자율
                oForm.DataSources.UserDataSources.Add("dj_eja1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("dj_eja1").Specific.DataBind.SetBound(true, "", "dj_eja1");

                oForm.DataSources.UserDataSources.Add("dj_eja2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("dj_eja2").Specific.DataBind.SetBound(true, "", "dj_eja2");

                oForm.DataSources.UserDataSources.Add("dj_eja3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("dj_eja3").Specific.DataBind.SetBound(true, "", "dj_eja3");

                // 계
                oForm.DataSources.UserDataSources.Add("dj_tamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_tamt1").Specific.DataBind.SetBound(true, "", "dj_tamt1");

                oForm.DataSources.UserDataSources.Add("dj_tamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_tamt2").Specific.DataBind.SetBound(true, "", "dj_tamt2");

                oForm.DataSources.UserDataSources.Add("dj_tamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_tamt3").Specific.DataBind.SetBound(true, "", "dj_tamt3");

                // 원리금
                oForm.DataSources.UserDataSources.Add("dj_wamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_wamt1").Specific.DataBind.SetBound(true, "", "dj_wamt1");

                oForm.DataSources.UserDataSources.Add("dj_wamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_wamt2").Specific.DataBind.SetBound(true, "", "dj_wamt2");

                oForm.DataSources.UserDataSources.Add("dj_wamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_wamt3").Specific.DataBind.SetBound(true, "", "dj_wamt3");

                // 이자
                oForm.DataSources.UserDataSources.Add("dj_eamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_eamt1").Specific.DataBind.SetBound(true, "", "dj_eamt1");

                oForm.DataSources.UserDataSources.Add("dj_eamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_eamt2").Specific.DataBind.SetBound(true, "", "dj_eamt2");

                oForm.DataSources.UserDataSources.Add("dj_eamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_eamt3").Specific.DataBind.SetBound(true, "", "dj_eamt3");

                // 공제금액
                oForm.DataSources.UserDataSources.Add("dj_gamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_gamt1").Specific.DataBind.SetBound(true, "", "dj_gamt1");

                oForm.DataSources.UserDataSources.Add("dj_gamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_gamt2").Specific.DataBind.SetBound(true, "", "dj_gamt2");

                oForm.DataSources.UserDataSources.Add("dj_gamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_gamt3").Specific.DataBind.SetBound(true, "", "dj_gamt3");

                // 임대차계약내용	
                // 임대인성명
                oForm.DataSources.UserDataSources.Add("ld_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ld_name1").Specific.DataBind.SetBound(true, "", "ld_name1");

                oForm.DataSources.UserDataSources.Add("ld_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ld_name2").Specific.DataBind.SetBound(true, "", "ld_name2");

                oForm.DataSources.UserDataSources.Add("ld_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ld_name3").Specific.DataBind.SetBound(true, "", "ld_name3");

                // 임대인주민등록번호
                oForm.DataSources.UserDataSources.Add("ld_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ld_jumin1").Specific.DataBind.SetBound(true, "", "ld_jumin1");

                oForm.DataSources.UserDataSources.Add("ld_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ld_jumin2").Specific.DataBind.SetBound(true, "", "ld_jumin2");

                oForm.DataSources.UserDataSources.Add("ld_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ld_jumin3").Specific.DataBind.SetBound(true, "", "ld_jumin3");

                // 주택유형
                oForm.DataSources.UserDataSources.Add("ld_hcode1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ld_hcode1").Specific.DataBind.SetBound(true, "", "ld_hcode1");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ld_hcode1").Specific, "Y");
                oForm.Items.Item("ld_hcode1").DisplayDesc = true;
                oForm.Items.Item("ld_hcode1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Add("ld_hcode2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ld_hcode2").Specific.DataBind.SetBound(true, "", "ld_hcode2");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ld_hcode2").Specific, "Y");
                oForm.Items.Item("ld_hcode2").DisplayDesc = true;
                oForm.Items.Item("ld_hcode2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Add("ld_hcode3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ld_hcode3").Specific.DataBind.SetBound(true, "", "ld_hcode3");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ld_hcode3").Specific, "Y");
                oForm.Items.Item("ld_hcode3").DisplayDesc = true;
                oForm.Items.Item("ld_hcode3").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 주택계약면적
                oForm.DataSources.UserDataSources.Add("ld_hm1", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ld_hm1").Specific.DataBind.SetBound(true, "", "ld_hm1");

                oForm.DataSources.UserDataSources.Add("ld_hm2", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ld_hm2").Specific.DataBind.SetBound(true, "", "ld_hm2");

                oForm.DataSources.UserDataSources.Add("ld_hm3", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ld_hm3").Specific.DataBind.SetBound(true, "", "ld_hm3");

                // 임대차계약서상주소
                oForm.DataSources.UserDataSources.Add("ld_addr1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ld_addr1").Specific.DataBind.SetBound(true, "", "ld_addr1");

                oForm.DataSources.UserDataSources.Add("ld_addr2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ld_addr2").Specific.DataBind.SetBound(true, "", "ld_addr2");

                oForm.DataSources.UserDataSources.Add("ld_addr3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ld_addr3").Specific.DataBind.SetBound(true, "", "ld_addr3");

                // 임대차계약기간
                oForm.DataSources.UserDataSources.Add("ld_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_fymd1").Specific.DataBind.SetBound(true, "", "ld_fymd1");

                oForm.DataSources.UserDataSources.Add("ld_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_fymd2").Specific.DataBind.SetBound(true, "", "ld_fymd2");

                oForm.DataSources.UserDataSources.Add("ld_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_fymd3").Specific.DataBind.SetBound(true, "", "ld_fymd3");

                oForm.DataSources.UserDataSources.Add("ld_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_tymd1").Specific.DataBind.SetBound(true, "", "ld_tymd1");

                oForm.DataSources.UserDataSources.Add("ld_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_tymd2").Specific.DataBind.SetBound(true, "", "ld_tymd2");

                oForm.DataSources.UserDataSources.Add("ld_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_tymd3").Specific.DataBind.SetBound(true, "", "ld_tymd3");

                //전세보증금
                oForm.DataSources.UserDataSources.Add("ld_bamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ld_bamt1").Specific.DataBind.SetBound(true, "", "ld_bamt1");

                oForm.DataSources.UserDataSources.Add("ld_bamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ld_bamt2").Specific.DataBind.SetBound(true, "", "ld_bamt2");

                oForm.DataSources.UserDataSources.Add("ld_bamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ld_bamt3").Specific.DataBind.SetBound(true, "", "ld_bamt3");

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
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY413_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oForm.EnableMenu("1282", true);  // 문서추가

                if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE)))
                {
                    oForm.Items.Item("Year").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);
                }
                //-----------------------------------------------
                oForm.DataSources.UserDataSources.Item("ws_name1").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_name2").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_name3").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_jumin1").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_jumin2").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_jumin3").Value = "";
                oForm.Items.Item("ws_hcode1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ws_hcode2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ws_hcode3").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("ws_hm1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_hm2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_hm3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_addr1").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_addr2").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_addr3").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_fymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_fymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_fymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_tymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_tymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_tymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_mamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_mamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_mamt3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_gamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_gamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ws_gamt3").Value = Convert.ToString(0);

                oForm.DataSources.UserDataSources.Item("dj_name1").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_name2").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_name3").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_jumin1").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_jumin2").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_jumin3").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_fymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_fymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_fymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_eja1").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_eja2").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_eja3").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_tamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_tamt3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_wamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_wamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_wamt3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_eamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_eamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_eamt3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_gamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_gamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("dj_gamt3").Value = Convert.ToString(0);

                oForm.DataSources.UserDataSources.Item("ld_name1").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_name2").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_name3").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_jumin1").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_jumin2").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_jumin3").Value = "";
                oForm.Items.Item("ld_hcode1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ld_hcode2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ld_hcode3").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("ld_hm1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ld_hm2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ld_hm3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ld_addr1").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_addr2").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_addr3").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_fymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_fymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_fymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_tymd1").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_tymd2").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_tymd3").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_bamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ld_bamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ld_bamt3").Value = Convert.ToString(0);

                //Key set
                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;
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
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY413);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
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
                            PH_PY413_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            string yyyy, Result = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY413_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY413_SAVE();
                        }
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {

                        yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY413_Delete();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            string yyyy = string.Empty;
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
            string seqn = string.Empty;
            string FullName = string.Empty;
            double amt = 0;
            double gamt = 0;
            double samt = 0;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                                sQry = "Select Code,";
                                sQry = sQry + " FullName = U_FullName,";
                                sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '1'";
                                sQry = sQry + " And U_Code = U_TeamCode),''),";
                                sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '2'";
                                sQry = sQry + " And U_Code = U_RspCode),''),";
                                sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '9'";
                                sQry = sQry + " And U_Code  = U_ClsCode";
                                sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "FullName":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                FullName = oForm.Items.Item("FullName").Specific.VALUE.Trim();

                                sQry = "Select Code,";
                                sQry = sQry + " FullName = U_FullName,";
                                sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '1'";
                                sQry = sQry + " And U_Code = U_TeamCode),''),";
                                sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '2'";
                                sQry = sQry + " And U_Code = U_RspCode),''),";
                                sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '9'";
                                sQry = sQry + " And U_Code  = U_ClsCode";
                                sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " And U_status <> '5'";    // 퇴사자 제외
                                sQry = sQry + " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "ws_mamt1":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                                //총급여액계산해서 5,500 이하는 15% 아니면 12%
                                sQry = "SELECT SUM(gwase) ";
                                sQry = sQry + "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "      Union All ";
                                sQry = sQry + "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry = sQry + "     ) g";

                                oRecordSet.DoQuery(sQry);
                                samt = oRecordSet.Fields.Item(0).Value;  // 총급여액(과세대상)

                                if (Convert.ToDouble(oForm.Items.Item("ws_mamt1").Specific.VALUE.Trim()) > 7500000)  // 한도 7백5십만원
                                {
                                    oForm.Items.Item("ws_mamt1").Specific.VALUE = 7500000;
                                }

                                gamt = 0;
                                if (samt <= 70000000)  // 7천이하자  10%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt1").Specific.VALUE.Trim()) * 0.1, 0);
                                }
                                if (samt <= 55000000)  // 5천5백이하자  12%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt1").Specific.VALUE.Trim()) * 0.12, 0);
                                }

                                if (gamt < 0)
                                {
                                    gamt = 0;
                                }

                                oForm.Items.Item("ws_gamt1").Specific.VALUE = gamt;
                                break;

                            case "ws_mamt2":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                                //총급여액계산해서 5,500 이하는 15% 아니면 12%
                                sQry = "SELECT SUM(gwase) ";
                                sQry = sQry + "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "      Union All ";
                                sQry = sQry + "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry = sQry + "     ) g";

                                oRecordSet.DoQuery(sQry);
                                samt = oRecordSet.Fields.Item(0).Value;  // 총급여액(과세대상)

                                if (Convert.ToDouble(oForm.Items.Item("ws_mamt2").Specific.VALUE.Trim()) > 7500000)  // 한도 7백5십만원
                                {
                                    oForm.Items.Item("ws_mamt2").Specific.VALUE = 7500000;
                                }

                                gamt = 0;
                                if (samt <= 70000000)  // 7천이하자  10%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt2").Specific.VALUE.Trim()) * 0.1, 0);
                                }
                                if (samt <= 55000000)  // 5천5백이하자  12%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt2").Specific.VALUE.Trim()) * 0.12, 0);
                                }

                                if (gamt < 0)
                                {
                                    gamt = 0;
                                }

                                oForm.Items.Item("ws_gamt2").Specific.VALUE = gamt;
                                break;

                            case "ws_mamt3":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                                //총급여액계산해서 5,500 이하는 15% 아니면 12%
                                sQry = "SELECT SUM(gwase) ";
                                sQry = sQry + "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "      Union All ";
                                sQry = sQry + "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry = sQry + "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry = sQry + "     ) g";

                                oRecordSet.DoQuery(sQry);
                                samt = oRecordSet.Fields.Item(0).Value;  // 총급여액(과세대상)

                                if (Convert.ToDouble(oForm.Items.Item("ws_mamt3").Specific.VALUE.Trim()) > 7500000)  // 한도 7백5십만원
                                {
                                    oForm.Items.Item("ws_mamt3").Specific.VALUE = 7500000;
                                }

                                gamt = 0;
                                if (samt <= 70000000)  // 7천이하자  10%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt3").Specific.VALUE.Trim()) * 0.1, 0);
                                }
                                if (samt <= 55000000)  // 5천5백이하자  12%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt3").Specific.VALUE.Trim()) * 0.12, 0);
                                }

                                if (gamt < 0)
                                {
                                    gamt = 0;
                                }

                                oForm.Items.Item("ws_gamt3").Specific.VALUE = gamt;
                                break;

                            case "dj_tamt1":
                                amt = 0;
                                gamt = 0;
                                amt = Convert.ToDouble(oForm.Items.Item("dj_tamt1").Specific.VALUE.Trim());
                                gamt = System.Math.Round(amt * 0.4, 0);
                                if (gamt > 3000000)
                                {
                                    oForm.Items.Item("dj_gamt1").Specific.VALUE = 3000000;
                                }
                                else
                                {
                                    oForm.Items.Item("dj_gamt1").Specific.VALUE = gamt;
                                }
                                break;

                            case "dj_tamt2":
                                amt = 0;
                                gamt = 0;
                                amt = Convert.ToDouble(oForm.Items.Item("dj_tamt2").Specific.VALUE.Trim());
                                gamt = System.Math.Round(amt * 0.4, 0);
                                if (gamt > 3000000)
                                {
                                    oForm.Items.Item("dj_gamt2").Specific.VALUE = 3000000;
                                }
                                else
                                {
                                    oForm.Items.Item("dj_gamt2").Specific.VALUE = gamt;
                                }
                                break;

                            case "dj_tamt3":
                                amt = 0;
                                gamt = 0;
                                amt = Convert.ToDouble(oForm.Items.Item("dj_tamt3").Specific.VALUE.Trim());
                                gamt = System.Math.Round(amt * 0.4, 0);
                                if (gamt > 3000000)
                                {
                                    oForm.Items.Item("dj_gamt3").Specific.VALUE = 3000000;
                                }
                                else
                                {
                                    oForm.Items.Item("dj_gamt3").Specific.VALUE = gamt;
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
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
        /// PH_PY413_DataFind
        /// </summary>
        private void PH_PY413_DataFind()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string Year = string.Empty;
            string MSTCOD = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Strings.Trim(Year)))
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(CLTCOD)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select * From [p_seoyhouse] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                oForm.DataSources.UserDataSources.Item("ws_name1").Value = oRecordSet.Fields.Item("ws_name1").Value;
                oForm.DataSources.UserDataSources.Item("ws_name2").Value = oRecordSet.Fields.Item("ws_name2").Value;
                oForm.DataSources.UserDataSources.Item("ws_name3").Value = oRecordSet.Fields.Item("ws_name3").Value;
                oForm.DataSources.UserDataSources.Item("ws_jumin1").Value = oRecordSet.Fields.Item("ws_jumin1").Value;
                oForm.DataSources.UserDataSources.Item("ws_jumin2").Value = oRecordSet.Fields.Item("ws_jumin2").Value;
                oForm.DataSources.UserDataSources.Item("ws_jumin3").Value = oRecordSet.Fields.Item("ws_jumin3").Value;
                oForm.Items.Item("ws_hcode1").Specific.Select(oRecordSet.Fields.Item("ws_hcode1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("ws_hcode2").Specific.Select(oRecordSet.Fields.Item("ws_hcode2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("ws_hcode3").Specific.Select(oRecordSet.Fields.Item("ws_hcode3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("ws_hm1").Value = oRecordSet.Fields.Item("ws_hm1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_hm2").Value = oRecordSet.Fields.Item("ws_hm2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_hm3").Value = oRecordSet.Fields.Item("ws_hm3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_addr1").Value = oRecordSet.Fields.Item("ws_addr1").Value;
                oForm.DataSources.UserDataSources.Item("ws_addr2").Value = oRecordSet.Fields.Item("ws_addr2").Value;
                oForm.DataSources.UserDataSources.Item("ws_addr3").Value = oRecordSet.Fields.Item("ws_addr3").Value;
                oForm.DataSources.UserDataSources.Item("ws_fymd1").Value = oRecordSet.Fields.Item("ws_fymd1").Value;
                oForm.DataSources.UserDataSources.Item("ws_fymd2").Value = oRecordSet.Fields.Item("ws_fymd2").Value;
                oForm.DataSources.UserDataSources.Item("ws_fymd3").Value = oRecordSet.Fields.Item("ws_fymd3").Value;
                oForm.DataSources.UserDataSources.Item("ws_tymd1").Value = oRecordSet.Fields.Item("ws_tymd1").Value;
                oForm.DataSources.UserDataSources.Item("ws_tymd2").Value = oRecordSet.Fields.Item("ws_tymd2").Value;
                oForm.DataSources.UserDataSources.Item("ws_tymd3").Value = oRecordSet.Fields.Item("ws_tymd3").Value;
                oForm.DataSources.UserDataSources.Item("ws_mamt1").Value = oRecordSet.Fields.Item("ws_mamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_mamt2").Value = oRecordSet.Fields.Item("ws_mamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_mamt3").Value = oRecordSet.Fields.Item("ws_mamt3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_gamt1").Value = oRecordSet.Fields.Item("ws_gamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_gamt2").Value = oRecordSet.Fields.Item("ws_gamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ws_gamt3").Value = oRecordSet.Fields.Item("ws_gamt3").Value.ToString();

                oForm.DataSources.UserDataSources.Item("dj_name1").Value = oRecordSet.Fields.Item("dj_name1").Value;
                oForm.DataSources.UserDataSources.Item("dj_name2").Value = oRecordSet.Fields.Item("dj_name2").Value;
                oForm.DataSources.UserDataSources.Item("dj_name3").Value = oRecordSet.Fields.Item("dj_name3").Value;
                oForm.DataSources.UserDataSources.Item("dj_jumin1").Value = oRecordSet.Fields.Item("dj_jumin1").Value;
                oForm.DataSources.UserDataSources.Item("dj_jumin2").Value = oRecordSet.Fields.Item("dj_jumin2").Value;
                oForm.DataSources.UserDataSources.Item("dj_jumin3").Value = oRecordSet.Fields.Item("dj_jumin3").Value;
                oForm.DataSources.UserDataSources.Item("dj_fymd1").Value = oRecordSet.Fields.Item("dj_fymd1").Value;
                oForm.DataSources.UserDataSources.Item("dj_fymd2").Value = oRecordSet.Fields.Item("dj_fymd2").Value;
                oForm.DataSources.UserDataSources.Item("dj_fymd3").Value = oRecordSet.Fields.Item("dj_fymd3").Value;
                oForm.DataSources.UserDataSources.Item("dj_tymd1").Value = oRecordSet.Fields.Item("dj_tymd1").Value;
                oForm.DataSources.UserDataSources.Item("dj_tymd2").Value = oRecordSet.Fields.Item("dj_tymd2").Value;
                oForm.DataSources.UserDataSources.Item("dj_tymd3").Value = oRecordSet.Fields.Item("dj_tymd3").Value;
                oForm.DataSources.UserDataSources.Item("dj_eja1").Value = oRecordSet.Fields.Item("dj_eja1").Value;
                oForm.DataSources.UserDataSources.Item("dj_eja2").Value = oRecordSet.Fields.Item("dj_eja2").Value;
                oForm.DataSources.UserDataSources.Item("dj_eja3").Value = oRecordSet.Fields.Item("dj_eja3").Value;
                oForm.DataSources.UserDataSources.Item("dj_tamt1").Value = oRecordSet.Fields.Item("dj_tamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_tamt2").Value = oRecordSet.Fields.Item("dj_tamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_tamt3").Value = oRecordSet.Fields.Item("dj_tamt3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_wamt1").Value = oRecordSet.Fields.Item("dj_wamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_wamt2").Value = oRecordSet.Fields.Item("dj_wamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_wamt3").Value = oRecordSet.Fields.Item("dj_wamt3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_eamt1").Value = oRecordSet.Fields.Item("dj_eamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_eamt2").Value = oRecordSet.Fields.Item("dj_eamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_eamt3").Value = oRecordSet.Fields.Item("dj_eamt3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_gamt1").Value = oRecordSet.Fields.Item("dj_gamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_gamt2").Value = oRecordSet.Fields.Item("dj_gamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("dj_gamt3").Value = oRecordSet.Fields.Item("dj_gamt3").Value.ToString();

                oForm.DataSources.UserDataSources.Item("ld_name1").Value = oRecordSet.Fields.Item("ld_name1").Value;
                oForm.DataSources.UserDataSources.Item("ld_name2").Value = oRecordSet.Fields.Item("ld_name2").Value;
                oForm.DataSources.UserDataSources.Item("ld_name3").Value = oRecordSet.Fields.Item("ld_name3").Value;
                oForm.DataSources.UserDataSources.Item("ld_jumin1").Value = oRecordSet.Fields.Item("ld_jumin1").Value;
                oForm.DataSources.UserDataSources.Item("ld_jumin2").Value = oRecordSet.Fields.Item("ld_jumin2").Value;
                oForm.DataSources.UserDataSources.Item("ld_jumin3").Value = oRecordSet.Fields.Item("ld_jumin3").Value;
                oForm.Items.Item("ld_hcode1").Specific.Select(oRecordSet.Fields.Item("ld_hcode1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("ld_hcode2").Specific.Select(oRecordSet.Fields.Item("ld_hcode2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("ld_hcode3").Specific.Select(oRecordSet.Fields.Item("ld_hcode3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("ld_hm1").Value = oRecordSet.Fields.Item("ld_hm1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ld_hm2").Value = oRecordSet.Fields.Item("ld_hm2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ld_hm3").Value = oRecordSet.Fields.Item("ld_hm3").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ld_addr1").Value = oRecordSet.Fields.Item("ld_addr1").Value;
                oForm.DataSources.UserDataSources.Item("ld_addr2").Value = oRecordSet.Fields.Item("ld_addr2").Value;
                oForm.DataSources.UserDataSources.Item("ld_addr3").Value = oRecordSet.Fields.Item("ld_addr3").Value;
                oForm.DataSources.UserDataSources.Item("ld_fymd1").Value = oRecordSet.Fields.Item("ld_fymd1").Value;
                oForm.DataSources.UserDataSources.Item("ld_fymd2").Value = oRecordSet.Fields.Item("ld_fymd2").Value;
                oForm.DataSources.UserDataSources.Item("ld_fymd3").Value = oRecordSet.Fields.Item("ld_fymd3").Value;
                oForm.DataSources.UserDataSources.Item("ld_tymd1").Value = oRecordSet.Fields.Item("ld_tymd1").Value;
                oForm.DataSources.UserDataSources.Item("ld_tymd2").Value = oRecordSet.Fields.Item("ld_tymd2").Value;
                oForm.DataSources.UserDataSources.Item("ld_tymd3").Value = oRecordSet.Fields.Item("ld_tymd3").Value;
                oForm.DataSources.UserDataSources.Item("ld_bamt1").Value = oRecordSet.Fields.Item("ld_bamt1").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ld_bamt2").Value = oRecordSet.Fields.Item("ld_bamt2").Value.ToString();
                oForm.DataSources.UserDataSources.Item("ld_bamt3").Value = oRecordSet.Fields.Item("ld_bamt3").Value.ToString();

                oForm.ActiveItem = "ws_name1";

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
        /// PH_PY413_SAVE
        /// </summary>
        private void PH_PY413_SAVE()
        {
            // 데이타 저장
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string saup, yyyy, sabun, ws_name1, ws_name2, ws_name3, ws_jumin1, ws_jumin2, ws_jumin3 = string.Empty;
            string ws_hcode1, ws_hcode2, ws_hcode3, ws_addr1, ws_addr2, ws_addr3, ws_fymd1, ws_fymd2, ws_fymd3 = string.Empty;
            string ws_tymd1, ws_tymd2, ws_tymd3, dj_name1, dj_name2, dj_name3, dj_jumin1, dj_jumin2, dj_jumin3 = string.Empty;
            string dj_fymd1, dj_fymd2, dj_fymd3, dj_tymd1, dj_tymd2, dj_tymd3, dj_eja1, dj_eja2, dj_eja3 = string.Empty;
            string ld_name1, ld_name2, ld_name3, ld_jumin1, ld_jumin2, ld_jumin3, ld_hcode1, ld_hcode2, ld_hcode3 = string.Empty;
            string ld_addr1, ld_addr2, ld_addr3, ld_fymd1, ld_fymd2, ld_fymd3, ld_tymd1, ld_tymd2, ld_tymd3 = string.Empty;

            double ws_hm1, ws_hm2, ws_hm3, ws_mamt1, ws_mamt2, ws_mamt3, ws_gamt1 , ws_gamt2, ws_gamt3 = 0;
            double dj_tamt1, dj_tamt2, dj_tamt3, dj_wamt1, dj_wamt2, dj_wamt3, dj_eamt1, dj_eamt2, dj_eamt3 = 0;
            double dj_gamt1, dj_gamt2, dj_gamt3, ld_hm1, ld_hm2, ld_hm3, ld_bamt1, ld_bamt2, ld_bamt3 = 0;

            try
            {
                oForm.Freeze(true);

                saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                yyyy = oForm.Items.Item("Year").Specific.VALUE;
                sabun = Strings.Trim(Conversion.Str(oForm.Items.Item("MSTCOD").Specific.VALUE));
                ws_name1 = oForm.Items.Item("ws_name1").Specific.VALUE;
                ws_name2 = oForm.Items.Item("ws_name2").Specific.VALUE;
                ws_name3 = oForm.Items.Item("ws_name3").Specific.VALUE;
                ws_jumin1 = oForm.Items.Item("ws_jumin1").Specific.VALUE;
                ws_jumin2 = oForm.Items.Item("ws_jumin2").Specific.VALUE;
                ws_jumin3 = oForm.Items.Item("ws_jumin3").Specific.VALUE;
                ws_hcode1 = oForm.Items.Item("ws_hcode1").Specific.VALUE;
                ws_hcode2 = oForm.Items.Item("ws_hcode2").Specific.VALUE;
                ws_hcode3 = oForm.Items.Item("ws_hcode3").Specific.VALUE;
                ws_hm1 = Convert.ToDouble(oForm.Items.Item("ws_hm1").Specific.VALUE);
                ws_hm2 = Convert.ToDouble(oForm.Items.Item("ws_hm2").Specific.VALUE);
                ws_hm3 = Convert.ToDouble(oForm.Items.Item("ws_hm3").Specific.VALUE);
                ws_addr1 = oForm.Items.Item("ws_addr1").Specific.VALUE;
                ws_addr2 = oForm.Items.Item("ws_addr2").Specific.VALUE;
                ws_addr3 = oForm.Items.Item("ws_addr3").Specific.VALUE;
                ws_fymd1 = oForm.Items.Item("ws_fymd1").Specific.VALUE;
                ws_fymd2 = oForm.Items.Item("ws_fymd2").Specific.VALUE;
                ws_fymd3 = oForm.Items.Item("ws_fymd3").Specific.VALUE;
                ws_tymd1 = oForm.Items.Item("ws_tymd1").Specific.VALUE;
                ws_tymd2 = oForm.Items.Item("ws_tymd2").Specific.VALUE;
                ws_tymd3 = oForm.Items.Item("ws_tymd3").Specific.VALUE;
                ws_mamt1 = Convert.ToDouble(oForm.Items.Item("ws_mamt1").Specific.VALUE);
                ws_mamt2 = Convert.ToDouble(oForm.Items.Item("ws_mamt2").Specific.VALUE);
                ws_mamt3 = Convert.ToDouble(oForm.Items.Item("ws_mamt3").Specific.VALUE);
                ws_gamt1 = Convert.ToDouble(oForm.Items.Item("ws_gamt1").Specific.VALUE);
                ws_gamt2 = Convert.ToDouble(oForm.Items.Item("ws_gamt2").Specific.VALUE);
                ws_gamt3 = Convert.ToDouble(oForm.Items.Item("ws_gamt3").Specific.VALUE);
                dj_name1 = oForm.Items.Item("dj_name1").Specific.VALUE;
                dj_name2 = oForm.Items.Item("dj_name2").Specific.VALUE;
                dj_name3 = oForm.Items.Item("dj_name3").Specific.VALUE;
                dj_jumin1 = oForm.Items.Item("dj_jumin1").Specific.VALUE;
                dj_jumin2 = oForm.Items.Item("dj_jumin2").Specific.VALUE;
                dj_jumin3 = oForm.Items.Item("dj_jumin3").Specific.VALUE;
                dj_fymd1 = oForm.Items.Item("dj_fymd1").Specific.VALUE;
                dj_fymd2 = oForm.Items.Item("dj_fymd2").Specific.VALUE;
                dj_fymd3 = oForm.Items.Item("dj_fymd3").Specific.VALUE;
                dj_tymd1 = oForm.Items.Item("dj_tymd1").Specific.VALUE;
                dj_tymd2 = oForm.Items.Item("dj_tymd2").Specific.VALUE;
                dj_tymd3 = oForm.Items.Item("dj_tymd3").Specific.VALUE;
                dj_eja1 = oForm.Items.Item("dj_eja1").Specific.VALUE;
                dj_eja2 = oForm.Items.Item("dj_eja2").Specific.VALUE;
                dj_eja3 = oForm.Items.Item("dj_eja3").Specific.VALUE;
                dj_tamt1 = Convert.ToDouble(oForm.Items.Item("dj_tamt1").Specific.VALUE);
                dj_tamt2 = Convert.ToDouble(oForm.Items.Item("dj_tamt2").Specific.VALUE);
                dj_tamt3 = Convert.ToDouble(oForm.Items.Item("dj_tamt3").Specific.VALUE);
                dj_wamt1 = Convert.ToDouble(oForm.Items.Item("dj_wamt1").Specific.VALUE);
                dj_wamt2 = Convert.ToDouble(oForm.Items.Item("dj_wamt2").Specific.VALUE);
                dj_wamt3 = Convert.ToDouble(oForm.Items.Item("dj_wamt3").Specific.VALUE);
                dj_eamt1 = Convert.ToDouble(oForm.Items.Item("dj_eamt1").Specific.VALUE);
                dj_eamt2 = Convert.ToDouble(oForm.Items.Item("dj_eamt2").Specific.VALUE);
                dj_eamt3 = Convert.ToDouble(oForm.Items.Item("dj_eamt3").Specific.VALUE);
                dj_gamt1 = Convert.ToDouble(oForm.Items.Item("dj_gamt1").Specific.VALUE);
                dj_gamt2 = Convert.ToDouble(oForm.Items.Item("dj_gamt2").Specific.VALUE);
                dj_gamt3 = Convert.ToDouble(oForm.Items.Item("dj_gamt3").Specific.VALUE);
                ld_name1 = oForm.Items.Item("ld_name1").Specific.VALUE;
                ld_name2 = oForm.Items.Item("ld_name2").Specific.VALUE;
                ld_name3 = oForm.Items.Item("ld_name3").Specific.VALUE;
                ld_jumin1 = oForm.Items.Item("ld_jumin1").Specific.VALUE;
                ld_jumin2 = oForm.Items.Item("ld_jumin2").Specific.VALUE;
                ld_jumin3 = oForm.Items.Item("ld_jumin3").Specific.VALUE;
                ld_hcode1 = oForm.Items.Item("ld_hcode1").Specific.VALUE;
                ld_hcode2 = oForm.Items.Item("ld_hcode2").Specific.VALUE;
                ld_hcode3 = oForm.Items.Item("ld_hcode3").Specific.VALUE;
                ld_hm1 = Convert.ToDouble(oForm.Items.Item("ld_hm1").Specific.VALUE);
                ld_hm2 = Convert.ToDouble(oForm.Items.Item("ld_hm2").Specific.VALUE);
                ld_hm3 = Convert.ToDouble(oForm.Items.Item("ld_hm3").Specific.VALUE);
                ld_addr1 = oForm.Items.Item("ld_addr1").Specific.VALUE;
                ld_addr2 = oForm.Items.Item("ld_addr2").Specific.VALUE;
                ld_addr3 = oForm.Items.Item("ld_addr3").Specific.VALUE;
                ld_fymd1 = oForm.Items.Item("ld_fymd1").Specific.VALUE;
                ld_fymd2 = oForm.Items.Item("ld_fymd2").Specific.VALUE;
                ld_fymd3 = oForm.Items.Item("ld_fymd3").Specific.VALUE;
                ld_tymd1 = oForm.Items.Item("ld_tymd1").Specific.VALUE;
                ld_tymd2 = oForm.Items.Item("ld_tymd2").Specific.VALUE;
                ld_tymd3 = oForm.Items.Item("ld_tymd3").Specific.VALUE;
                ld_bamt1 = Convert.ToDouble(oForm.Items.Item("ld_bamt1").Specific.VALUE);
                ld_bamt2 = Convert.ToDouble(oForm.Items.Item("ld_bamt2").Specific.VALUE);
                ld_bamt3 = Convert.ToDouble(oForm.Items.Item("ld_bamt3").Specific.VALUE);

                if (string.IsNullOrEmpty(Strings.Trim(yyyy)))
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(saup)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(sabun)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    ////갱신

                    sQry = "Update [p_seoyhouse] set ";
                    sQry = sQry + "ws_name1 = '" + ws_name1 + "',";
                    sQry = sQry + "ws_name2 = '" + ws_name2 + "',";
                    sQry = sQry + "ws_name3 = '" + ws_name3 + "',";
                    sQry = sQry + "ws_jumin1 = '" + ws_jumin1 + "',";
                    sQry = sQry + "ws_jumin2 = '" + ws_jumin2 + "',";
                    sQry = sQry + "ws_jumin3 = '" + ws_jumin3 + "',";
                    sQry = sQry + "ws_hcode1 = '" + ws_hcode1 + "',";
                    sQry = sQry + "ws_hcode2 = '" + ws_hcode2 + "',";
                    sQry = sQry + "ws_hcode3 = '" + ws_hcode3 + "',";
                    sQry = sQry + "ws_hm1 = '" + ws_hm1 + "',";
                    sQry = sQry + "ws_hm2 = '" + ws_hm2 + "',";
                    sQry = sQry + "ws_hm3 = '" + ws_hm3 + "',";
                    sQry = sQry + "ws_addr1 = '" + ws_addr1 + "',";
                    sQry = sQry + "ws_addr2 = '" + ws_addr2 + "',";
                    sQry = sQry + "ws_addr3 = '" + ws_addr3 + "',";
                    sQry = sQry + "ws_fymd1 = '" + ws_fymd1 + "',";
                    sQry = sQry + "ws_fymd2 = '" + ws_fymd2 + "',";
                    sQry = sQry + "ws_fymd3 = '" + ws_fymd3 + "',";
                    sQry = sQry + "ws_tymd1 = '" + ws_tymd1 + "',";
                    sQry = sQry + "ws_tymd2 = '" + ws_tymd2 + "',";
                    sQry = sQry + "ws_tymd3 = '" + ws_tymd3 + "',";
                    sQry = sQry + "ws_mamt1 = " + ws_mamt1 + ",";
                    sQry = sQry + "ws_mamt2 = " + ws_mamt2 + ",";
                    sQry = sQry + "ws_mamt3 = " + ws_mamt3 + ",";
                    sQry = sQry + "ws_gamt1 = " + ws_gamt1 + ",";
                    sQry = sQry + "ws_gamt2 = " + ws_gamt2 + ",";
                    sQry = sQry + "ws_gamt3 = " + ws_gamt3 + ",";

                    sQry = sQry + "dj_name1 = '" + dj_name1 + "',";
                    sQry = sQry + "dj_name2 = '" + dj_name2 + "',";
                    sQry = sQry + "dj_name3 = '" + dj_name3 + "',";
                    sQry = sQry + "dj_jumin1 = '" + dj_jumin1 + "',";
                    sQry = sQry + "dj_jumin2 = '" + dj_jumin2 + "',";
                    sQry = sQry + "dj_jumin3 = '" + dj_jumin3 + "',";
                    sQry = sQry + "dj_fymd1 = '" + dj_fymd1 + "',";
                    sQry = sQry + "dj_fymd2 = '" + dj_fymd2 + "',";
                    sQry = sQry + "dj_fymd3 = '" + dj_fymd3 + "',";
                    sQry = sQry + "dj_tymd1 = '" + dj_tymd1 + "',";
                    sQry = sQry + "dj_tymd2 = '" + dj_tymd2 + "',";
                    sQry = sQry + "dj_tymd3 = '" + dj_tymd3 + "',";
                    sQry = sQry + "dj_eja1 = '" + dj_eja1 + "',";
                    sQry = sQry + "dj_eja2 = '" + dj_eja2 + "',";
                    sQry = sQry + "dj_eja3 = '" + dj_eja3 + "',";
                    sQry = sQry + "dj_tamt1 = " + dj_tamt1 + ",";
                    sQry = sQry + "dj_tamt2 = " + dj_tamt2 + ",";
                    sQry = sQry + "dj_tamt3 = " + dj_tamt3 + ",";
                    sQry = sQry + "dj_wamt1 = " + dj_wamt1 + ",";
                    sQry = sQry + "dj_wamt2 = " + dj_wamt2 + ",";
                    sQry = sQry + "dj_wamt3 = " + dj_wamt3 + ",";
                    sQry = sQry + "dj_eamt1 = " + dj_eamt1 + ",";
                    sQry = sQry + "dj_eamt2 = " + dj_eamt2 + ",";
                    sQry = sQry + "dj_eamt3 = " + dj_eamt3 + ",";
                    sQry = sQry + "dj_gamt1 = " + dj_gamt1 + ",";
                    sQry = sQry + "dj_gamt2 = " + dj_gamt2 + ",";
                    sQry = sQry + "dj_gamt3 = " + dj_gamt3 + ",";

                    sQry = sQry + "ld_name1 = '" + ld_name1 + "',";
                    sQry = sQry + "ld_name2 = '" + ld_name2 + "',";
                    sQry = sQry + "ld_name3 = '" + ld_name3 + "',";
                    sQry = sQry + "ld_jumin1 = '" + ld_jumin1 + "',";
                    sQry = sQry + "ld_jumin2 = '" + ld_jumin2 + "',";
                    sQry = sQry + "ld_jumin3 = '" + ld_jumin3 + "',";
                    sQry = sQry + "ld_hcode1 = '" + ld_hcode1 + "',";
                    sQry = sQry + "ld_hcode2 = '" + ld_hcode2 + "',";
                    sQry = sQry + "ld_hcode3 = '" + ld_hcode3 + "',";
                    sQry = sQry + "ld_hm1 = '" + ld_hm1 + "',";
                    sQry = sQry + "ld_hm2 = '" + ld_hm2 + "',";
                    sQry = sQry + "ld_hm3 = '" + ld_hm3 + "',";
                    sQry = sQry + "ld_addr1 = '" + ld_addr1 + "',";
                    sQry = sQry + "ld_addr2 = '" + ld_addr2 + "',";
                    sQry = sQry + "ld_addr3 = '" + ld_addr3 + "',";
                    sQry = sQry + "ld_fymd1 = '" + ld_fymd1 + "',";
                    sQry = sQry + "ld_fymd2 = '" + ld_fymd2 + "',";
                    sQry = sQry + "ld_fymd3 = '" + ld_fymd3 + "',";
                    sQry = sQry + "ld_tymd1 = '" + ld_tymd1 + "',";
                    sQry = sQry + "ld_tymd2 = '" + ld_tymd2 + "',";
                    sQry = sQry + "ld_tymd3 = '" + ld_tymd3 + "',";
                    sQry = sQry + "ld_bamt1 = " + ld_bamt1 + ",";
                    sQry = sQry + "ld_bamt2 = " + ld_bamt2 + ",";
                    sQry = sQry + "ld_bamt3 = " + ld_bamt3 + "";

                    sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY413_DataFind();

                }
                else
                {

                    ////신규
                    sQry = "INSERT INTO [p_seoyhouse]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "ws_name1, ";
                    sQry = sQry + "ws_name2, ";
                    sQry = sQry + "ws_name3, ";
                    sQry = sQry + "ws_jumin1, ";
                    sQry = sQry + "ws_jumin2, ";
                    sQry = sQry + "ws_jumin3, ";
                    sQry = sQry + "ws_hcode1, ";
                    sQry = sQry + "ws_hcode2, ";
                    sQry = sQry + "ws_hcode3, ";
                    sQry = sQry + "ws_hm1, ";
                    sQry = sQry + "ws_hm2, ";
                    sQry = sQry + "ws_hm3, ";
                    sQry = sQry + "ws_addr1, ";
                    sQry = sQry + "ws_addr2, ";
                    sQry = sQry + "ws_addr3, ";
                    sQry = sQry + "ws_fymd1, ";
                    sQry = sQry + "ws_fymd2, ";
                    sQry = sQry + "ws_fymd3, ";
                    sQry = sQry + "ws_tymd1, ";
                    sQry = sQry + "ws_tymd2, ";
                    sQry = sQry + "ws_tymd3, ";
                    sQry = sQry + "ws_mamt1, ";
                    sQry = sQry + "ws_mamt2, ";
                    sQry = sQry + "ws_mamt3, ";
                    sQry = sQry + "ws_gamt1, ";
                    sQry = sQry + "ws_gamt2, ";
                    sQry = sQry + "ws_gamt3, ";

                    sQry = sQry + "dj_name1, ";
                    sQry = sQry + "dj_name2, ";
                    sQry = sQry + "dj_name3, ";
                    sQry = sQry + "dj_jumin1, ";
                    sQry = sQry + "dj_jumin2, ";
                    sQry = sQry + "dj_jumin3, ";
                    sQry = sQry + "dj_fymd1, ";
                    sQry = sQry + "dj_fymd2, ";
                    sQry = sQry + "dj_fymd3, ";
                    sQry = sQry + "dj_tymd1, ";
                    sQry = sQry + "dj_tymd2, ";
                    sQry = sQry + "dj_tymd3, ";
                    sQry = sQry + "dj_eja1, ";
                    sQry = sQry + "dj_eja2, ";
                    sQry = sQry + "dj_eja3, ";
                    sQry = sQry + "dj_tamt1, ";
                    sQry = sQry + "dj_tamt2, ";
                    sQry = sQry + "dj_tamt3, ";
                    sQry = sQry + "dj_wamt1, ";
                    sQry = sQry + "dj_wamt2, ";
                    sQry = sQry + "dj_wamt3, ";
                    sQry = sQry + "dj_eamt1, ";
                    sQry = sQry + "dj_eamt2, ";
                    sQry = sQry + "dj_eamt3, ";
                    sQry = sQry + "dj_gamt1, ";
                    sQry = sQry + "dj_gamt2, ";
                    sQry = sQry + "dj_gamt3, ";

                    sQry = sQry + "ld_name1, ";
                    sQry = sQry + "ld_name2, ";
                    sQry = sQry + "ld_name3, ";
                    sQry = sQry + "ld_jumin1, ";
                    sQry = sQry + "ld_jumin2, ";
                    sQry = sQry + "ld_jumin3, ";
                    sQry = sQry + "ld_hcode1, ";
                    sQry = sQry + "ld_hcode2, ";
                    sQry = sQry + "ld_hcode3, ";
                    sQry = sQry + "ld_hm1, ";
                    sQry = sQry + "ld_hm2, ";
                    sQry = sQry + "ld_hm3, ";
                    sQry = sQry + "ld_addr1, ";
                    sQry = sQry + "ld_addr2, ";
                    sQry = sQry + "ld_addr3, ";
                    sQry = sQry + "ld_fymd1, ";
                    sQry = sQry + "ld_fymd2, ";
                    sQry = sQry + "ld_fymd3, ";
                    sQry = sQry + "ld_tymd1, ";
                    sQry = sQry + "ld_tymd2, ";
                    sQry = sQry + "ld_tymd3, ";
                    sQry = sQry + "ld_bamt1, ";
                    sQry = sQry + "ld_bamt2, ";
                    sQry = sQry + "ld_bamt3 ";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + saup + "',";
                    sQry = sQry + "'" + yyyy + "',";
                    sQry = sQry + "'" + sabun + "',";
                    sQry = sQry + "'" + ws_name1 + "',";
                    sQry = sQry + "'" + ws_name2 + "',";
                    sQry = sQry + "'" + ws_name3 + "',";
                    sQry = sQry + "'" + ws_jumin1 + "',";
                    sQry = sQry + "'" + ws_jumin2 + "',";
                    sQry = sQry + "'" + ws_jumin3 + "',";
                    sQry = sQry + "'" + ws_hcode1 + "',";
                    sQry = sQry + "'" + ws_hcode2 + "',";
                    sQry = sQry + "'" + ws_hcode3 + "',";
                    sQry = sQry + ws_hm1 + ",";
                    sQry = sQry + ws_hm2 + ",";
                    sQry = sQry + ws_hm3 + ",";
                    sQry = sQry + "'" + ws_addr1 + "',";
                    sQry = sQry + "'" + ws_addr2 + "',";
                    sQry = sQry + "'" + ws_addr3 + "',";
                    sQry = sQry + "'" + ws_fymd1 + "',";
                    sQry = sQry + "'" + ws_fymd2 + "',";
                    sQry = sQry + "'" + ws_fymd3 + "',";
                    sQry = sQry + "'" + ws_tymd1 + "',";
                    sQry = sQry + "'" + ws_tymd2 + "',";
                    sQry = sQry + "'" + ws_tymd3 + "',";
                    sQry = sQry + ws_mamt1 + ",";
                    sQry = sQry + ws_mamt2 + ",";
                    sQry = sQry + ws_mamt3 + ",";
                    sQry = sQry + ws_gamt1 + ",";
                    sQry = sQry + ws_gamt2 + ",";
                    sQry = sQry + ws_gamt3 + ",";

                    sQry = sQry + "'" + dj_name1 + "',";
                    sQry = sQry + "'" + dj_name2 + "',";
                    sQry = sQry + "'" + dj_name3 + "',";
                    sQry = sQry + "'" + dj_jumin1 + "',";
                    sQry = sQry + "'" + dj_jumin2 + "',";
                    sQry = sQry + "'" + dj_jumin3 + "',";
                    sQry = sQry + "'" + dj_fymd1 + "',";
                    sQry = sQry + "'" + dj_fymd2 + "',";
                    sQry = sQry + "'" + dj_fymd3 + "',";
                    sQry = sQry + "'" + dj_tymd1 + "',";
                    sQry = sQry + "'" + dj_tymd2 + "',";
                    sQry = sQry + "'" + dj_tymd3 + "',";
                    sQry = sQry + "'" + dj_eja1 + "',";
                    sQry = sQry + "'" + dj_eja2 + "',";
                    sQry = sQry + "'" + dj_eja3 + "',";
                    sQry = sQry + dj_tamt1 + ",";
                    sQry = sQry + dj_tamt2 + ",";
                    sQry = sQry + dj_tamt3 + ",";
                    sQry = sQry + dj_wamt1 + ",";
                    sQry = sQry + dj_wamt2 + ",";
                    sQry = sQry + dj_wamt3 + ",";
                    sQry = sQry + dj_eamt1 + ",";
                    sQry = sQry + dj_eamt2 + ",";
                    sQry = sQry + dj_eamt3 + ",";
                    sQry = sQry + dj_gamt1 + ",";
                    sQry = sQry + dj_gamt2 + ",";
                    sQry = sQry + dj_gamt3 + ",";

                    sQry = sQry + "'" + ld_name1 + "',";
                    sQry = sQry + "'" + ld_name2 + "',";
                    sQry = sQry + "'" + ld_name3 + "',";
                    sQry = sQry + "'" + ld_jumin1 + "',";
                    sQry = sQry + "'" + ld_jumin2 + "',";
                    sQry = sQry + "'" + ld_jumin3 + "',";
                    sQry = sQry + "'" + ld_hcode1 + "',";
                    sQry = sQry + "'" + ld_hcode2 + "',";
                    sQry = sQry + "'" + ld_hcode3 + "',";
                    sQry = sQry + ld_hm1 + ",";
                    sQry = sQry + ld_hm2 + ",";
                    sQry = sQry + ld_hm3 + ",";
                    sQry = sQry + "'" + ld_addr1 + "',";
                    sQry = sQry + "'" + ld_addr2 + "',";
                    sQry = sQry + "'" + ld_addr3 + "',";
                    sQry = sQry + "'" + ld_fymd1 + "',";
                    sQry = sQry + "'" + ld_fymd2 + "',";
                    sQry = sQry + "'" + ld_fymd3 + "',";
                    sQry = sQry + "'" + ld_tymd1 + "',";
                    sQry = sQry + "'" + ld_tymd2 + "',";
                    sQry = sQry + "'" + ld_tymd3 + "',";
                    sQry = sQry + ld_bamt1 + ",";
                    sQry = sQry + ld_bamt2 + ",";
                    sQry = sQry + ld_bamt3;
                    sQry = sQry + " ) ";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY413_DataFind();
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
        /// PH_PY413_Delete
        /// </summary>
        private void PH_PY413_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string saup, yyyy, sabun, seqn = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    sQry = "Delete From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY413_DataFind();
                    oForm.ActiveItem = "MSTCOD";
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    //    PSH_Globals.SBO_Application.MessageBox("급여계산 된 자료는 삭제할 수 없습니다.");
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
    }
}
