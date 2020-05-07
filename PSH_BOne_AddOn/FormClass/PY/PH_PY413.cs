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
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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



//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY413
//	{
//////********************************************************************************
//////  File           : PH_PY413.cls
//////  Module         : 인사관리 > 연말정산관리
//////  Desc           : 월세액.주택임차차입금자료 등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Grid oGrid1;
//		public SAPbouiCOM.DataTable oDS_PH_PY413A;


//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY413.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY413_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY413");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "Code"

//			oForm.PaneLevel = 1;
//			oForm.Freeze(true);
//			PH_PY413_CreateItems();
//			PH_PY413_FormItemEnabled();
//			PH_PY413_EnableMenus();
//			//    Call PH_PY413_SetDocument(oFromDocEntry01)
//			//    Call PH_PY413_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY413_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;
//			string CLTCOD = null;

//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbouiCOM.OptionBtn optBtn = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//Set oGrid1 = oForm.Items("Grid01").Specific

//			oForm.DataSources.DataTables.Add("PH_PY413");

//			//oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY413")
//			oDS_PH_PY413A = oForm.DataSources.DataTables.Item("PH_PY413");


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			////사업장

//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			//    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
//			//    oCombo.Select CLTCOD, psk_ByValue
//			//    oCombo.Select 0, psk_Index
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			////년도
//			oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

//			////사번
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
//			////성명
//			oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

//			////임대인성명
//			oForm.DataSources.UserDataSources.Add("ws_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_name1").Specific.DataBind.SetBound(true, "", "ws_name1");

//			oForm.DataSources.UserDataSources.Add("ws_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_name2").Specific.DataBind.SetBound(true, "", "ws_name2");

//			oForm.DataSources.UserDataSources.Add("ws_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_name3").Specific.DataBind.SetBound(true, "", "ws_name3");

//			////임대인주민등록번호
//			oForm.DataSources.UserDataSources.Add("ws_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_jumin1").Specific.DataBind.SetBound(true, "", "ws_jumin1");

//			oForm.DataSources.UserDataSources.Add("ws_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_jumin2").Specific.DataBind.SetBound(true, "", "ws_jumin2");

//			oForm.DataSources.UserDataSources.Add("ws_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_jumin3").Specific.DataBind.SetBound(true, "", "ws_jumin3");

//			////주택유형
//			oCombo = oForm.Items.Item("ws_hcode1").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ws_hcode1").DisplayDesc = true;

//			oCombo = oForm.Items.Item("ws_hcode2").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ws_hcode2").DisplayDesc = true;

//			oCombo = oForm.Items.Item("ws_hcode3").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ws_hcode3").DisplayDesc = true;

//			////주택계약면적
//			oForm.DataSources.UserDataSources.Add("ws_hm1", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_hm1").Specific.DataBind.SetBound(true, "", "ws_hm1");

//			oForm.DataSources.UserDataSources.Add("ws_hm2", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_hm2").Specific.DataBind.SetBound(true, "", "ws_hm2");

//			oForm.DataSources.UserDataSources.Add("ws_hm3", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_hm3").Specific.DataBind.SetBound(true, "", "ws_hm3");

//			////임대차계약서상주소
//			oForm.DataSources.UserDataSources.Add("ws_addr1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_addr1").Specific.DataBind.SetBound(true, "", "ws_addr1");

//			oForm.DataSources.UserDataSources.Add("ws_addr2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_addr2").Specific.DataBind.SetBound(true, "", "ws_addr2");

//			oForm.DataSources.UserDataSources.Add("ws_addr3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_addr3").Specific.DataBind.SetBound(true, "", "ws_addr3");

//			////임대차계약기간
//			oForm.DataSources.UserDataSources.Add("ws_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_fymd1").Specific.DataBind.SetBound(true, "", "ws_fymd1");

//			oForm.DataSources.UserDataSources.Add("ws_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_fymd2").Specific.DataBind.SetBound(true, "", "ws_fymd2");

//			oForm.DataSources.UserDataSources.Add("ws_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_fymd3").Specific.DataBind.SetBound(true, "", "ws_fymd3");

//			oForm.DataSources.UserDataSources.Add("ws_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_tymd1").Specific.DataBind.SetBound(true, "", "ws_tymd1");

//			oForm.DataSources.UserDataSources.Add("ws_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_tymd2").Specific.DataBind.SetBound(true, "", "ws_tymd2");

//			oForm.DataSources.UserDataSources.Add("ws_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_tymd3").Specific.DataBind.SetBound(true, "", "ws_tymd3");

//			////월세액
//			oForm.DataSources.UserDataSources.Add("ws_mamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_mamt1").Specific.DataBind.SetBound(true, "", "ws_mamt1");

//			oForm.DataSources.UserDataSources.Add("ws_mamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_mamt2").Specific.DataBind.SetBound(true, "", "ws_mamt2");

//			oForm.DataSources.UserDataSources.Add("ws_mamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_mamt3").Specific.DataBind.SetBound(true, "", "ws_mamt3");

//			////공제금액
//			oForm.DataSources.UserDataSources.Add("ws_gamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_gamt1").Specific.DataBind.SetBound(true, "", "ws_gamt1");

//			oForm.DataSources.UserDataSources.Add("ws_gamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_gamt2").Specific.DataBind.SetBound(true, "", "ws_gamt2");

//			oForm.DataSources.UserDataSources.Add("ws_gamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ws_gamt3").Specific.DataBind.SetBound(true, "", "ws_gamt3");

//			////대주
//			oForm.DataSources.UserDataSources.Add("dj_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_name1").Specific.DataBind.SetBound(true, "", "dj_name1");

//			oForm.DataSources.UserDataSources.Add("dj_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_name2").Specific.DataBind.SetBound(true, "", "dj_name2");

//			oForm.DataSources.UserDataSources.Add("dj_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_name3").Specific.DataBind.SetBound(true, "", "dj_name3");

//			////대주 주민등록번호
//			oForm.DataSources.UserDataSources.Add("dj_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_jumin1").Specific.DataBind.SetBound(true, "", "dj_jumin1");

//			oForm.DataSources.UserDataSources.Add("dj_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_jumin2").Specific.DataBind.SetBound(true, "", "dj_jumin2");

//			oForm.DataSources.UserDataSources.Add("dj_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_jumin3").Specific.DataBind.SetBound(true, "", "dj_jumin3");

//			////금전소비대차계약기간
//			oForm.DataSources.UserDataSources.Add("dj_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_fymd1").Specific.DataBind.SetBound(true, "", "dj_fymd1");

//			oForm.DataSources.UserDataSources.Add("dj_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_fymd2").Specific.DataBind.SetBound(true, "", "dj_fymd2");

//			oForm.DataSources.UserDataSources.Add("dj_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_fymd3").Specific.DataBind.SetBound(true, "", "dj_fymd3");

//			oForm.DataSources.UserDataSources.Add("dj_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tymd1").Specific.DataBind.SetBound(true, "", "dj_tymd1");

//			oForm.DataSources.UserDataSources.Add("dj_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tymd2").Specific.DataBind.SetBound(true, "", "dj_tymd2");

//			oForm.DataSources.UserDataSources.Add("dj_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tymd3").Specific.DataBind.SetBound(true, "", "dj_tymd3");

//			////차임급이자율
//			oForm.DataSources.UserDataSources.Add("dj_eja1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eja1").Specific.DataBind.SetBound(true, "", "dj_eja1");

//			oForm.DataSources.UserDataSources.Add("dj_eja2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eja2").Specific.DataBind.SetBound(true, "", "dj_eja2");

//			oForm.DataSources.UserDataSources.Add("dj_eja3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eja3").Specific.DataBind.SetBound(true, "", "dj_eja3");

//			////계
//			oForm.DataSources.UserDataSources.Add("dj_tamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tamt1").Specific.DataBind.SetBound(true, "", "dj_tamt1");

//			oForm.DataSources.UserDataSources.Add("dj_tamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tamt2").Specific.DataBind.SetBound(true, "", "dj_tamt2");

//			oForm.DataSources.UserDataSources.Add("dj_tamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_tamt3").Specific.DataBind.SetBound(true, "", "dj_tamt3");

//			////원리금
//			oForm.DataSources.UserDataSources.Add("dj_wamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_wamt1").Specific.DataBind.SetBound(true, "", "dj_wamt1");

//			oForm.DataSources.UserDataSources.Add("dj_wamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_wamt2").Specific.DataBind.SetBound(true, "", "dj_wamt2");

//			oForm.DataSources.UserDataSources.Add("dj_wamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_wamt3").Specific.DataBind.SetBound(true, "", "dj_wamt3");

//			////이자
//			oForm.DataSources.UserDataSources.Add("dj_eamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eamt1").Specific.DataBind.SetBound(true, "", "dj_eamt1");

//			oForm.DataSources.UserDataSources.Add("dj_eamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eamt2").Specific.DataBind.SetBound(true, "", "dj_eamt2");

//			oForm.DataSources.UserDataSources.Add("dj_eamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_eamt3").Specific.DataBind.SetBound(true, "", "dj_eamt3");

//			////공제금액
//			oForm.DataSources.UserDataSources.Add("dj_gamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_gamt1").Specific.DataBind.SetBound(true, "", "dj_gamt1");

//			oForm.DataSources.UserDataSources.Add("dj_gamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_gamt2").Specific.DataBind.SetBound(true, "", "dj_gamt2");

//			oForm.DataSources.UserDataSources.Add("dj_gamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("dj_gamt3").Specific.DataBind.SetBound(true, "", "dj_gamt3");

//			////임대인성명
//			oForm.DataSources.UserDataSources.Add("ld_name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_name1").Specific.DataBind.SetBound(true, "", "ld_name1");

//			oForm.DataSources.UserDataSources.Add("ld_name2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_name2").Specific.DataBind.SetBound(true, "", "ld_name2");

//			oForm.DataSources.UserDataSources.Add("ld_name3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_name3").Specific.DataBind.SetBound(true, "", "ld_name3");

//			////임대인주민등록번호
//			oForm.DataSources.UserDataSources.Add("ld_jumin1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_jumin1").Specific.DataBind.SetBound(true, "", "ld_jumin1");

//			oForm.DataSources.UserDataSources.Add("ld_jumin2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_jumin2").Specific.DataBind.SetBound(true, "", "ld_jumin2");

//			oForm.DataSources.UserDataSources.Add("ld_jumin3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_jumin3").Specific.DataBind.SetBound(true, "", "ld_jumin3");

//			////주택유형
//			oCombo = oForm.Items.Item("ld_hcode1").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ld_hcode1").DisplayDesc = true;

//			oCombo = oForm.Items.Item("ld_hcode2").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ld_hcode2").DisplayDesc = true;

//			oCombo = oForm.Items.Item("ld_hcode3").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("ld_hcode3").DisplayDesc = true;

//			////주택계약면적
//			oForm.DataSources.UserDataSources.Add("ld_hm1", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_hm1").Specific.DataBind.SetBound(true, "", "ld_hm1");

//			oForm.DataSources.UserDataSources.Add("ld_hm2", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_hm2").Specific.DataBind.SetBound(true, "", "ld_hm2");

//			oForm.DataSources.UserDataSources.Add("ld_hm3", SAPbouiCOM.BoDataType.dt_QUANTITY);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_hm3").Specific.DataBind.SetBound(true, "", "ld_hm3");

//			////임대차계약서상주소
//			oForm.DataSources.UserDataSources.Add("ld_addr1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_addr1").Specific.DataBind.SetBound(true, "", "ld_addr1");

//			oForm.DataSources.UserDataSources.Add("ld_addr2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_addr2").Specific.DataBind.SetBound(true, "", "ld_addr2");

//			oForm.DataSources.UserDataSources.Add("ld_addr3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_addr3").Specific.DataBind.SetBound(true, "", "ld_addr3");

//			////임대차계약기간
//			oForm.DataSources.UserDataSources.Add("ld_fymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_fymd1").Specific.DataBind.SetBound(true, "", "ld_fymd1");

//			oForm.DataSources.UserDataSources.Add("ld_fymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_fymd2").Specific.DataBind.SetBound(true, "", "ld_fymd2");

//			oForm.DataSources.UserDataSources.Add("ld_fymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_fymd3").Specific.DataBind.SetBound(true, "", "ld_fymd3");

//			oForm.DataSources.UserDataSources.Add("ld_tymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_tymd1").Specific.DataBind.SetBound(true, "", "ld_tymd1");

//			oForm.DataSources.UserDataSources.Add("ld_tymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_tymd2").Specific.DataBind.SetBound(true, "", "ld_tymd2");

//			oForm.DataSources.UserDataSources.Add("ld_tymd3", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_tymd3").Specific.DataBind.SetBound(true, "", "ld_tymd3");

//			////전세보증금
//			oForm.DataSources.UserDataSources.Add("ld_bamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_bamt1").Specific.DataBind.SetBound(true, "", "ld_bamt1");

//			oForm.DataSources.UserDataSources.Add("ld_bamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_bamt2").Specific.DataBind.SetBound(true, "", "ld_bamt2");

//			oForm.DataSources.UserDataSources.Add("ld_bamt3", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ld_bamt3").Specific.DataBind.SetBound(true, "", "ld_bamt3");


//			oForm.Update();

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY413_CreateItems_Error:

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY413_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", false);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", false);
//			////행삭제

//			return;
//			PH_PY413_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY413_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY413_FormItemEnabled();
//				//        Call PH_PY413_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY413_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY413_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY413_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			int i = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			string CLTCOD = null;
//			string sPosDate = null;


//			 // ERROR: Not supported in C#: OnErrorStatement

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//				//UPGRADE_WARNING: oForm.Items(Year).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Year").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;
//				//oForm.Items("MSTCOD").Specific.VALUE = ""
//				//oForm.Items("FullName").Specific.VALUE = ""
//				//oForm.Items("TeamName").Specific.VALUE = ""
//				//oForm.Items("RspName").Specific.VALUE = ""
//				//oForm.Items("ClsName").Specific.VALUE = ""
//				//-----------------------------------------------
//				oForm.DataSources.UserDataSources.Item("ws_name1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_name2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_name3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_jumin1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_jumin2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_jumin3").Value = "";
//				oCombo = oForm.Items.Item("ws_hcode1").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oCombo = oForm.Items.Item("ws_hcode2").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oCombo = oForm.Items.Item("ws_hcode3").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oForm.DataSources.UserDataSources.Item("ws_hm1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_hm2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_hm3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_addr1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_addr2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_addr3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_fymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_fymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_fymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_tymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_tymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_tymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ws_mamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_mamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_mamt3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_gamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_gamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ws_gamt3").Value = Convert.ToString(0);

//				oForm.DataSources.UserDataSources.Item("dj_name1").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_name2").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_name3").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_jumin1").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_jumin2").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_jumin3").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_fymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_fymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_fymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_tymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_tymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_tymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_eja1").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_eja2").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_eja3").Value = "";
//				oForm.DataSources.UserDataSources.Item("dj_tamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_tamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_tamt3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_wamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_wamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_wamt3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_eamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_eamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_eamt3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_gamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_gamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("dj_gamt3").Value = Convert.ToString(0);

//				oForm.DataSources.UserDataSources.Item("ld_name1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_name2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_name3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_jumin1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_jumin2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_jumin3").Value = "";
//				oCombo = oForm.Items.Item("ld_hcode1").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oCombo = oForm.Items.Item("ld_hcode2").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oCombo = oForm.Items.Item("ld_hcode3").Specific;
//				oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				oForm.DataSources.UserDataSources.Item("ld_hm1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ld_hm2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ld_hm3").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ld_addr1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_addr2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_addr3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_fymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_fymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_fymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_tymd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_tymd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_tymd3").Value = "";
//				oForm.DataSources.UserDataSources.Item("ld_bamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ld_bamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ld_bamt3").Value = Convert.ToString(0);

//				//-----------------------------------------------

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가


//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}

//			////Key set
//			oForm.Items.Item("CLTCOD").Enabled = true;
//			oForm.Items.Item("Year").Enabled = true;
//			oForm.Items.Item("MSTCOD").Enabled = true;

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY413_FormItemEnabled_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			string tSex = null;
//			string tBrith = null;
//			//UPGRADE_NOTE: Day이(가) Day_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string Day_Renamed = null;
//			string ActCode = null;
//			string CLTCOD = null;
//			string MSTCOD = null;
//			string FullName = null;
//			int Amt = 0;
//			int gamt = 0;
//			string YY = null;
//			string Result = null;
//			string yyyy = null;

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (PH_PY413_DataValidCheck() == false) {
//								BubbleEvent = false;
//							}
//						}

//						if (pval.ItemUID == "Btn_ret") {
//							PH_PY413_MTX01();
//						}

//						if (pval.ItemUID == "Btn01") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							yyyy = oForm.Items.Item("Year").Specific.VALUE;
//							sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
//							oRecordSet.DoQuery(sQry);

//							//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							Result = oRecordSet.Fields.Item(0).Value;
//							if (Result != "Y") {
//								MDC_Globals.Sbo_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
//							}
//							if (Result == "Y") {
//								PH_PY413_SAVE();
//							}
//						}


//						if (pval.ItemUID == "Btn_del") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							yyyy = oForm.Items.Item("Year").Specific.VALUE;
//							sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
//							oRecordSet.DoQuery(sQry);

//							//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							Result = oRecordSet.Fields.Item(0).Value;
//							if (Result != "Y") {
//								MDC_Globals.Sbo_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
//							}
//							if (Result == "Y") {
//								PH_PY413_Delete();
//								PH_PY413_FormItemEnabled();
//							}
//						}
//						//                If oForm.Mode = fm_FIND_MODE Then
//						//                    If pval.ItemUID = "Btn01" Then
//						//                        Sbo_Application.ActivateMenuItem ("7425")
//						//                        BubbleEvent = False
//						//                    End If
//						//
//						//                End If
//					} else if (pval.BeforeAction == false) {
//						switch (pval.ItemUID) {
//							case "1":
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY413_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY413_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY413_FormItemEnabled();
//									}
//								}
//								break;
//							//
//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true) {
//						if (pval.CharPressed == 9) {
//							if (pval.ItemUID == "MSTCOD") {
//								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							}
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							////사업장(헤더)
//							if (pval.ItemUID == "SCLTCOD") {

//							}

//						}
//					}

//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row >= 0) {
//									switch (pval.ItemUID) {
//										case "Grid01":
//											//Call oMat1.SelectRow(pval.Row, True, False)
//											PH_PY413_MTX02(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//											break;
//									}

//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//					} else {

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					//            Call oForm.Freeze(True)
//					if (pval.BeforeAction == true) {
//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							switch (pval.ItemUID) {
//								case "MSTCOD":
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

//									sQry = "Select Code,";
//									sQry = sQry + " FullName = U_FullName,";
//									sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '1'";
//									sQry = sQry + " And U_Code = U_TeamCode),''),";
//									sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '2'";
//									sQry = sQry + " And U_Code = U_RspCode),''),";
//									sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '9'";
//									sQry = sQry + " And U_Code  = U_ClsCode";
//									sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
//									sQry = sQry + " From [@PH_PY001A]";
//									sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
//									sQry = sQry + " and Code = '" + MSTCOD + "'";

//									oRecordSet.DoQuery(sQry);

//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
//									//                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
//									//UPGRADE_WARNING: oForm.Items(TeamName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
//									//UPGRADE_WARNING: oForm.Items(RspName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
//									//UPGRADE_WARNING: oForm.Items(ClsName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;

//									////조회실행
//									PH_PY413_MTX01();
//									break;
//								case "FullName":
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									FullName = oForm.Items.Item("FullName").Specific.VALUE;

//									sQry = "Select Code,";
//									sQry = sQry + " FullName = U_FullName,";
//									sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '1'";
//									sQry = sQry + " And U_Code = U_TeamCode),''),";
//									sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '2'";
//									sQry = sQry + " And U_Code = U_RspCode),''),";
//									sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '9'";
//									sQry = sQry + " And U_Code  = U_ClsCode";
//									sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
//									sQry = sQry + " From [@PH_PY001A]";
//									sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
//									sQry = sQry + " And U_status <> '5'";
//									//퇴사자 제외
//									sQry = sQry + " and U_FullName = '" + FullName + "'";

//									oRecordSet.DoQuery(sQry);

//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
//									//                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
//									//UPGRADE_WARNING: oForm.Items(TeamName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
//									//UPGRADE_WARNING: oForm.Items(RspName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
//									//UPGRADE_WARNING: oForm.Items(ClsName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;

//									////조회실행
//									PH_PY413_MTX01();
//									break;

//								case "ws_mamt1":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = System.Math.Round(oForm.Items.Item("ws_mamt1").Specific.VALUE * 0.1, 0);
//									gamt = Amt;
//									if (gamt > 750000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt1").Specific.VALUE = 750000;
//										//14년한도7,500,000 에 세액공제10%(75만원)
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt1").Specific.VALUE = gamt;
//									}

//									//UPGRADE_WARNING: oForm.Items(ws_gamt1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("ws_gamt1").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt1").Specific.VALUE = 0;
//									}
//									break;

//								case "ws_mamt2":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = System.Math.Round(oForm.Items.Item("ws_mamt2").Specific.VALUE * 0.1, 0);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									gamt = oForm.Items.Item("ws_gamt1").Specific.VALUE;

//									if (gamt + Amt > 750000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt2").Specific.VALUE = 750000 - gamt;
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt2").Specific.VALUE = Amt;
//									}

//									//UPGRADE_WARNING: oForm.Items(ws_gamt2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("ws_gamt2").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt2").Specific.VALUE = 0;
//									}
//									break;

//								case "ws_mamt3":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = System.Math.Round(oForm.Items.Item("ws_mamt3").Specific.VALUE * 0.1, 0);
//									//UPGRADE_WARNING: oForm.Items(ws_gamt2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									gamt = oForm.Items.Item("ws_gamt1").Specific.VALUE + oForm.Items.Item("ws_gamt2").Specific.VALUE;
//									if (gamt + Amt > 750000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt3").Specific.VALUE = 750000 - gamt;
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt3").Specific.VALUE = Amt;
//									}

//									//UPGRADE_WARNING: oForm.Items(ws_gamt3).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("ws_gamt3").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("ws_gamt3").Specific.VALUE = 0;
//									}
//									break;

//								case "dj_tamt1":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = oForm.Items.Item("dj_tamt1").Specific.VALUE;
//									gamt = System.Math.Round(Amt * 0.4, 0);
//									if (gamt > 3000000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt1").Specific.VALUE = 3000000;
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt1").Specific.VALUE = gamt;
//									}

//									//UPGRADE_WARNING: oForm.Items(dj_gamt1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("dj_gamt1").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt1").Specific.VALUE = 0;
//									}
//									break;

//								case "dj_tamt2":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = oForm.Items.Item("dj_tamt2").Specific.VALUE;
//									gamt = System.Math.Round(Amt * 0.4, 0);
//									if (gamt > 3000000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt2").Specific.VALUE = 3000000;
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt2").Specific.VALUE = gamt;
//									}

//									//UPGRADE_WARNING: oForm.Items(dj_gamt2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("dj_gamt2").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt2").Specific.VALUE = 0;
//									}
//									break;

//								case "dj_tamt3":
//									Amt = 0;
//									gamt = 0;
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									Amt = oForm.Items.Item("dj_tamt3").Specific.VALUE;
//									gamt = System.Math.Round(Amt * 0.4, 0);
//									if (gamt > 3000000) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt3").Specific.VALUE = 3000000;
//									} else {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt3").Specific.VALUE = gamt;
//									}

//									//UPGRADE_WARNING: oForm.Items(dj_gamt3).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("dj_gamt3").Specific.VALUE < 0) {
//										//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oForm.Items.Item("dj_gamt3").Specific.VALUE = 0;
//									}
//									break;

//							}

//						}
//					}
//					break;
//				//            Call oForm.Freeze(False)
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						//                oMat1.LoadFromDataSource
//						//                Call PH_PY413_AddMatrixRow

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY413A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY413A = null;

//						//                Set oMat1 = Nothing
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY413A", "Code")
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;


//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;
//					//                Call PH_PY413_FormItemEnabled
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY413_FormItemEnabled();
//						break;
//					//                Call PH_PY413_AddMatrixRow
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY413_FormItemEnabled();
//						//                Call PH_PY413_AddMatrixRow
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY413_FormItemEnabled();
//						break;
//					//                Call PH_PY413_AddMatrixRow
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY413_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			int i = 0;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}

//			}
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Raise_FormDataEvent_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat01":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void PH_PY413_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY413'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY413_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY413_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			int j = 0;

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY413_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY413_MTX01()
//		{

//			////DATA조회

//			int i = 0;
//			string sQry = null;
//			int iRow = 0;
//			SAPbouiCOM.ComboBox oCombo = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("MSTCOD").Specific.VALUE;

//			if (string.IsNullOrEmpty(Strings.Trim(Param01))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_MTX01_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Param02))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_MTX01_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Param03))) {
//				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_MTX01_Exit;
//			}



//			sQry = " Select * From [p_seoyhouse] Where saup = '" + Param01 + "' And yyyy = '" + Param02 + "' And sabun = '" + Param03 + "'";
//			oRecordSet.DoQuery(sQry);

//			//oForm.Items("ws_name1").Specific.VALUE = oRecordSet.Fields("ws_name1").VALUE

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_name1").Value = oRecordSet.Fields.Item("ws_name1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_name2").Value = oRecordSet.Fields.Item("ws_name2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_name3").Value = oRecordSet.Fields.Item("ws_name3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_jumin1").Value = oRecordSet.Fields.Item("ws_jumin1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_jumin2").Value = oRecordSet.Fields.Item("ws_jumin2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_jumin3").Value = oRecordSet.Fields.Item("ws_jumin3").Value;
//			oCombo = oForm.Items.Item("ws_hcode1").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ws_hcode1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("ws_hcode2").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ws_hcode2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("ws_hcode3").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ws_hcode3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_hm1").Value = oRecordSet.Fields.Item("ws_hm1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_hm2").Value = oRecordSet.Fields.Item("ws_hm2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_hm3").Value = oRecordSet.Fields.Item("ws_hm3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_addr1").Value = oRecordSet.Fields.Item("ws_addr1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_addr2").Value = oRecordSet.Fields.Item("ws_addr2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_addr3").Value = oRecordSet.Fields.Item("ws_addr3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_fymd1").Value = oRecordSet.Fields.Item("ws_fymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_fymd2").Value = oRecordSet.Fields.Item("ws_fymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_fymd3").Value = oRecordSet.Fields.Item("ws_fymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_tymd1").Value = oRecordSet.Fields.Item("ws_tymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_tymd2").Value = oRecordSet.Fields.Item("ws_tymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_tymd3").Value = oRecordSet.Fields.Item("ws_tymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_mamt1").Value = oRecordSet.Fields.Item("ws_mamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_mamt2").Value = oRecordSet.Fields.Item("ws_mamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_mamt3").Value = oRecordSet.Fields.Item("ws_mamt3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_gamt1").Value = oRecordSet.Fields.Item("ws_gamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_gamt2").Value = oRecordSet.Fields.Item("ws_gamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ws_gamt3").Value = oRecordSet.Fields.Item("ws_gamt3").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_name1").Value = oRecordSet.Fields.Item("dj_name1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_name2").Value = oRecordSet.Fields.Item("dj_name2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_name3").Value = oRecordSet.Fields.Item("dj_name3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_jumin1").Value = oRecordSet.Fields.Item("dj_jumin1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_jumin2").Value = oRecordSet.Fields.Item("dj_jumin2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_jumin3").Value = oRecordSet.Fields.Item("dj_jumin3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_fymd1").Value = oRecordSet.Fields.Item("dj_fymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_fymd2").Value = oRecordSet.Fields.Item("dj_fymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_fymd3").Value = oRecordSet.Fields.Item("dj_fymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tymd1").Value = oRecordSet.Fields.Item("dj_tymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tymd2").Value = oRecordSet.Fields.Item("dj_tymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tymd3").Value = oRecordSet.Fields.Item("dj_tymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eja1").Value = oRecordSet.Fields.Item("dj_eja1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eja2").Value = oRecordSet.Fields.Item("dj_eja2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eja3").Value = oRecordSet.Fields.Item("dj_eja3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tamt1").Value = oRecordSet.Fields.Item("dj_tamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tamt2").Value = oRecordSet.Fields.Item("dj_tamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_tamt3").Value = oRecordSet.Fields.Item("dj_tamt3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_wamt1").Value = oRecordSet.Fields.Item("dj_wamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_wamt2").Value = oRecordSet.Fields.Item("dj_wamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_wamt3").Value = oRecordSet.Fields.Item("dj_wamt3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eamt1").Value = oRecordSet.Fields.Item("dj_eamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eamt2").Value = oRecordSet.Fields.Item("dj_eamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_eamt3").Value = oRecordSet.Fields.Item("dj_eamt3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_gamt1").Value = oRecordSet.Fields.Item("dj_gamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_gamt2").Value = oRecordSet.Fields.Item("dj_gamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("dj_gamt3").Value = oRecordSet.Fields.Item("dj_gamt3").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_name1").Value = oRecordSet.Fields.Item("ld_name1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_name2").Value = oRecordSet.Fields.Item("ld_name2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_name3").Value = oRecordSet.Fields.Item("ld_name3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_jumin1").Value = oRecordSet.Fields.Item("ld_jumin1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_jumin2").Value = oRecordSet.Fields.Item("ld_jumin2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_jumin3").Value = oRecordSet.Fields.Item("ld_jumin3").Value;
//			oCombo = oForm.Items.Item("ld_hcode1").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ld_hcode1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("ld_hcode2").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ld_hcode2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("ld_hcode3").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("ld_hcode3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_hm1").Value = oRecordSet.Fields.Item("ld_hm1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_hm2").Value = oRecordSet.Fields.Item("ld_hm2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_hm3").Value = oRecordSet.Fields.Item("ld_hm3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_addr1").Value = oRecordSet.Fields.Item("ld_addr1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_addr2").Value = oRecordSet.Fields.Item("ld_addr2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_addr3").Value = oRecordSet.Fields.Item("ld_addr3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_fymd1").Value = oRecordSet.Fields.Item("ld_fymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_fymd2").Value = oRecordSet.Fields.Item("ld_fymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_fymd3").Value = oRecordSet.Fields.Item("ld_fymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_tymd1").Value = oRecordSet.Fields.Item("ld_tymd1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_tymd2").Value = oRecordSet.Fields.Item("ld_tymd2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_tymd3").Value = oRecordSet.Fields.Item("ld_tymd3").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_bamt1").Value = oRecordSet.Fields.Item("ld_bamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_bamt2").Value = oRecordSet.Fields.Item("ld_bamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ld_bamt3").Value = oRecordSet.Fields.Item("ld_bamt3").Value;


//			oForm.ActiveItem = "ws_name1";

//			oForm.Update();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY413_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY413_MTX01_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY413_MTX02(string oUID, ref int oRow = 0, ref string oCol = "")
//		{


//			//    '//그리드 자료를 head에 로드
//			//
//			//    Dim i       As Long
//			//    Dim sQry    As String
//			//    Dim sRow As Long
//			//
//			//    Dim Param01 As String
//			//    Dim Param02 As String
//			//    Dim Param03 As String
//			//    Dim Param04 As String
//			//    Dim Param05 As String
//			//    Dim Param06 As String
//			//    Dim Param07 As String
//			//
//			//    Dim oCombo      As SAPbouiCOM.ComboBox
//			//    Dim oRecordSet As SAPbobsCOM.Recordset
//			//
//			//    On Error GoTo PH_PY413_MTX02_Error
//			//
//			//    Call oForm.Freeze(True)
//			//    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
//			//
//			//    sRow = oRow
//			//
//			//
//			//    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
//			//    Param02 = oDS_PH_PY413A.Columns.Item("연도").Cells(oRow).VALUE
//			//    Param03 = oDS_PH_PY413A.Columns.Item("사번").Cells(oRow).VALUE
//			//    Param04 = oDS_PH_PY413A.Columns.Item("주민번호").Cells(oRow).VALUE
//			//    Param05 = oDS_PH_PY413A.Columns.Item("지급처상호").Cells(oRow).VALUE
//			//    Param06 = oDS_PH_PY413A.Columns.Item("지급일자").Cells(oRow).VALUE
//			//    Param07 = oDS_PH_PY413A.Columns.Item("사업자번호").Cells(oRow).VALUE
//			//
//			//
//			//
//			//    sQry = "EXEC PH_PY413_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "'"
//			//    Call oRecordSet.DoQuery(sQry)
//			//
//			//    If (oRecordSet.RecordCount = 0) Then
//			//
//			//        'oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY413A.Columns.Item("MSTCOD").Cells(oRow).VALUE
//			//        'oForm.Items("FullName").Specific.VALUE = oDS_PH_PY413A.Columns.Item("FullName").Cells(oRow).VALUE
//			//
//			//        oForm.Items("kname").Specific.VALUE = ""
//			//        oForm.Items("juminno").Specific.VALUE = ""
//			//        oForm.Items("custnm").Specific.VALUE = ""
//			//        oForm.Items("entno").Specific.VALUE = ""
//			//        oForm.Items("payymd").Specific.VALUE = ""
//			//
//			//        oForm.Items("medcex").Specific.VALUE = 0
//			//        oForm.Items("ntamt").Specific.VALUE = 0
//			//        oForm.Items("cont").Specific.VALUE = 0
//			//
//			//        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
//			//        GoTo PH_PY413_MTX02_Exit
//			//    End If
//			//
//			//    Set oCombo = oForm.Items("rel").Specific
//			//    oCombo.Select oRecordSet.Fields("rel").VALUE, psk_ByValue
//			//
//			//    oForm.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
//			//    oForm.DataSources.UserDataSources.Item("juminno").VALUE = oRecordSet.Fields("juminno").VALUE
//			//
//			//    Set oCombo = oForm.Items("empdiv").Specific
//			//    oCombo.Select oRecordSet.Fields("empdiv").VALUE, psk_ByValue
//			//
//			//    oForm.DataSources.UserDataSources.Item("custnm").VALUE = oRecordSet.Fields("custnm").VALUE
//			//    oForm.DataSources.UserDataSources.Item("entno").VALUE = oRecordSet.Fields("entno").VALUE
//			//    oForm.DataSources.UserDataSources.Item("payymd").VALUE = oRecordSet.Fields("payymd").VALUE
//			//
//			//    Set oCombo = oForm.Items("gubun").Specific
//			//    oCombo.Select oRecordSet.Fields("gubun").VALUE, psk_ByValue
//			//
//			//    oForm.DataSources.UserDataSources.Item("medcex").VALUE = oRecordSet.Fields("medcex").VALUE
//			//    oForm.DataSources.UserDataSources.Item("ntamt").VALUE = oRecordSet.Fields("ntamt").VALUE
//			//    oForm.DataSources.UserDataSources.Item("cont").VALUE = oRecordSet.Fields("cont").VALUE
//			//
//			//    Set oCombo = oForm.Items("olddiv").Specific
//			//    oCombo.Select oRecordSet.Fields("olddiv").VALUE, psk_ByValue
//			//
//			//    Set oCombo = oForm.Items("deform").Specific
//			//    oCombo.Select oRecordSet.Fields("deform").VALUE, psk_ByValue
//			//
//			//'    '//부서
//			//'    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
//			//'    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
//			//'    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
//			//
//			//    '//Key Disable
//			//    oForm.Items("CLTCOD").Enabled = False
//			//    oForm.Items("Year").Enabled = False
//			//    oForm.Items("MSTCOD").Enabled = False
//			//
//			//    oForm.Items("juminno").Enabled = False
//			//    oForm.Items("custnm").Enabled = False
//			//    oForm.Items("payymd").Enabled = False
//			//    oForm.Items("entno").Enabled = False
//			//
//			//
//			//    oForm.Update
//			//
//			//    Set oRecordSet = Nothing
//			//    Call oForm.Freeze(False)
//			//    Exit Sub
//			//PH_PY413_MTX02_Exit:
//			//    Set oRecordSet = Nothing
//			//    Call oForm.Freeze(False)
//			//
//			//    Exit Sub
//			//PH_PY413_MTX02_Error:
//			//    Set oRecordSet = Nothing
//			//    Call oForm.Freeze(False)
//			//    Sbo_Application.SetStatusBarMessage "PH_PY413_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
//		}

//		public bool PH_PY413_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY413A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY413A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY413_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY413_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY413_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//////행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent, ref SAPbouiCOM.Matrix oMat, ref SAPbouiCOM.DBDataSource DBData, ref string CheckField)
//		{

//			int i = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((oLastColRow > 0)) {
//				if (pval.BeforeAction == true) {

//				} else if (pval.BeforeAction == false) {
//					if (oMat.RowCount != oMat.VisualRowCount) {
//						oMat.FlushToDataSource();

//						while ((i <= DBData.Size - 1)) {
//							if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i))) {
//								DBData.RemoveRecord((i));
//								i = 0;
//							} else {
//								i = i + 1;
//							}
//						}

//						for (i = 0; i <= DBData.Size; i++) {
//							DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//						}

//						oMat.LoadFromDataSource();
//					}
//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY413_SAVE()
//		{

//			////데이타 저장

//			int i = 0;
//			string sQry = null;

//			string yyyy = null;
//			string saup = null;
//			string sabun = null;
//			string ws_addr2 = null;
//			string ws_jumin3 = null;
//			string ws_jumin1 = null;
//			string ws_name2 = null;
//			string ws_name1 = null;
//			string ws_name3 = null;
//			string ws_jumin2 = null;
//			string ws_addr1 = null;
//			string ws_addr3 = null;
//			string ws_tymd2 = null;
//			string ws_fymd3 = null;
//			string ws_fymd1 = null;
//			string ws_fymd2 = null;
//			string ws_tymd1 = null;
//			string ws_tymd3 = null;
//			object ws_gamt1 = null;
//			object ws_mamt2 = null;
//			object ws_mamt1 = null;
//			object ws_mamt3 = null;
//			object ws_gamt2 = null;
//			double ws_gamt3 = 0;
//			string ws_hcode2 = null;
//			string ws_hcode1 = null;
//			string ws_hcode3 = null;
//			object ws_hm1 = null;
//			object ws_hm2 = null;
//			double ws_hm3 = 0;

//			string dj_jumin2 = null;
//			string dj_name3 = null;
//			string dj_name1 = null;
//			string dj_name2 = null;
//			string dj_jumin1 = null;
//			string dj_jumin3 = null;
//			string dj_tymd2 = null;
//			string dj_fymd3 = null;
//			string dj_fymd1 = null;
//			string dj_fymd2 = null;
//			string dj_tymd1 = null;
//			string dj_tymd3 = null;
//			string dj_eja2 = null;
//			string dj_eja1 = null;
//			string dj_eja3 = null;
//			object dj_wamt1 = null;
//			object dj_tamt2 = null;
//			object dj_tamt1 = null;
//			object dj_tamt3 = null;
//			object dj_wamt2 = null;
//			double dj_wamt3 = 0;
//			object dj_gamt1 = null;
//			object dj_eamt2 = null;
//			object dj_eamt1 = null;
//			object dj_eamt3 = null;
//			object dj_gamt2 = null;
//			double dj_gamt3 = 0;

//			string ld_addr2 = null;
//			string ld_jumin3 = null;
//			string ld_jumin1 = null;
//			string ld_name2 = null;
//			string ld_name1 = null;
//			string ld_name3 = null;
//			string ld_jumin2 = null;
//			string ld_addr1 = null;
//			string ld_addr3 = null;
//			string ld_tymd2 = null;
//			string ld_fymd3 = null;
//			string ld_fymd1 = null;
//			string ld_fymd2 = null;
//			string ld_tymd1 = null;
//			string ld_tymd3 = null;
//			object ld_bamt1 = null;
//			object ld_bamt2 = null;
//			double ld_bamt3 = 0;
//			string ld_hcode2 = null;
//			string ld_hcode1 = null;
//			string ld_hcode3 = null;
//			object ld_hm1 = null;
//			object ld_hm2 = null;
//			double ld_hm3 = 0;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sabun = Strings.Trim(Conversion.Str(oForm.Items.Item("MSTCOD").Specific.VALUE));

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_name1 = oForm.Items.Item("ws_name1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_name2 = oForm.Items.Item("ws_name2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_name3 = oForm.Items.Item("ws_name3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_jumin1 = oForm.Items.Item("ws_jumin1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_jumin2 = oForm.Items.Item("ws_jumin2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_jumin3 = oForm.Items.Item("ws_jumin3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hcode1 = oForm.Items.Item("ws_hcode1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hcode2 = oForm.Items.Item("ws_hcode2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hcode3 = oForm.Items.Item("ws_hcode3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hm1 = oForm.Items.Item("ws_hm1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hm2 = oForm.Items.Item("ws_hm2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_hm3 = oForm.Items.Item("ws_hm3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_addr1 = oForm.Items.Item("ws_addr1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_addr2 = oForm.Items.Item("ws_addr2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_addr3 = oForm.Items.Item("ws_addr3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_fymd1 = oForm.Items.Item("ws_fymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_fymd2 = oForm.Items.Item("ws_fymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_fymd3 = oForm.Items.Item("ws_fymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_tymd1 = oForm.Items.Item("ws_tymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_tymd2 = oForm.Items.Item("ws_tymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_tymd3 = oForm.Items.Item("ws_tymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_mamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_mamt1 = oForm.Items.Item("ws_mamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_mamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_mamt2 = oForm.Items.Item("ws_mamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_mamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_mamt3 = oForm.Items.Item("ws_mamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_gamt1 = oForm.Items.Item("ws_gamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ws_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_gamt2 = oForm.Items.Item("ws_gamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ws_gamt3 = oForm.Items.Item("ws_gamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_name1 = oForm.Items.Item("dj_name1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_name2 = oForm.Items.Item("dj_name2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_name3 = oForm.Items.Item("dj_name3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_jumin1 = oForm.Items.Item("dj_jumin1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_jumin2 = oForm.Items.Item("dj_jumin2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_jumin3 = oForm.Items.Item("dj_jumin3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_fymd1 = oForm.Items.Item("dj_fymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_fymd2 = oForm.Items.Item("dj_fymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_fymd3 = oForm.Items.Item("dj_fymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tymd1 = oForm.Items.Item("dj_tymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tymd2 = oForm.Items.Item("dj_tymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tymd3 = oForm.Items.Item("dj_tymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eja1 = oForm.Items.Item("dj_eja1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eja2 = oForm.Items.Item("dj_eja2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eja3 = oForm.Items.Item("dj_eja3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_tamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tamt1 = oForm.Items.Item("dj_tamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_tamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tamt2 = oForm.Items.Item("dj_tamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_tamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_tamt3 = oForm.Items.Item("dj_tamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_wamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_wamt1 = oForm.Items.Item("dj_wamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_wamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_wamt2 = oForm.Items.Item("dj_wamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_wamt3 = oForm.Items.Item("dj_wamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_eamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eamt1 = oForm.Items.Item("dj_eamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_eamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eamt2 = oForm.Items.Item("dj_eamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_eamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_eamt3 = oForm.Items.Item("dj_eamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_gamt1 = oForm.Items.Item("dj_gamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: dj_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_gamt2 = oForm.Items.Item("dj_gamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			dj_gamt3 = oForm.Items.Item("dj_gamt3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_name1 = oForm.Items.Item("ld_name1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_name2 = oForm.Items.Item("ld_name2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_name3 = oForm.Items.Item("ld_name3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_jumin1 = oForm.Items.Item("ld_jumin1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_jumin2 = oForm.Items.Item("ld_jumin2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_jumin3 = oForm.Items.Item("ld_jumin3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hcode1 = oForm.Items.Item("ld_hcode1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hcode2 = oForm.Items.Item("ld_hcode2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hcode3 = oForm.Items.Item("ld_hcode3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ld_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hm1 = oForm.Items.Item("ld_hm1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ld_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hm2 = oForm.Items.Item("ld_hm2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_hm3 = oForm.Items.Item("ld_hm3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_addr1 = oForm.Items.Item("ld_addr1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_addr2 = oForm.Items.Item("ld_addr2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_addr3 = oForm.Items.Item("ld_addr3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_fymd1 = oForm.Items.Item("ld_fymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_fymd2 = oForm.Items.Item("ld_fymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_fymd3 = oForm.Items.Item("ld_fymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_tymd1 = oForm.Items.Item("ld_tymd1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_tymd2 = oForm.Items.Item("ld_tymd2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_tymd3 = oForm.Items.Item("ld_tymd3").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ld_bamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_bamt1 = oForm.Items.Item("ld_bamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ld_bamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_bamt2 = oForm.Items.Item("ld_bamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ld_bamt3 = oForm.Items.Item("ld_bamt3").Specific.VALUE;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if (string.IsNullOrEmpty(Strings.Trim(yyyy))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_SAVE_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(saup))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_SAVE_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(sabun))) {
//				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY413_SAVE_Exit;
//			}

//			sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.Fields.Item(0).Value > 0) {
//				////갱신

//				sQry = "Update [p_seoyhouse] set ";
//				sQry = sQry + "ws_name1 = '" + ws_name1 + "',";
//				sQry = sQry + "ws_name2 = '" + ws_name2 + "',";
//				sQry = sQry + "ws_name3 = '" + ws_name3 + "',";
//				sQry = sQry + "ws_jumin1 = '" + ws_jumin1 + "',";
//				sQry = sQry + "ws_jumin2 = '" + ws_jumin2 + "',";
//				sQry = sQry + "ws_jumin3 = '" + ws_jumin3 + "',";
//				sQry = sQry + "ws_hcode1 = '" + ws_hcode1 + "',";
//				sQry = sQry + "ws_hcode2 = '" + ws_hcode2 + "',";
//				sQry = sQry + "ws_hcode3 = '" + ws_hcode3 + "',";
//				//UPGRADE_WARNING: ws_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_hm1 = '" + ws_hm1 + "',";
//				//UPGRADE_WARNING: ws_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_hm2 = '" + ws_hm2 + "',";
//				sQry = sQry + "ws_hm3 = '" + ws_hm3 + "',";
//				sQry = sQry + "ws_addr1 = '" + ws_addr1 + "',";
//				sQry = sQry + "ws_addr2 = '" + ws_addr2 + "',";
//				sQry = sQry + "ws_addr3 = '" + ws_addr3 + "',";
//				sQry = sQry + "ws_fymd1 = '" + ws_fymd1 + "',";
//				sQry = sQry + "ws_fymd2 = '" + ws_fymd2 + "',";
//				sQry = sQry + "ws_fymd3 = '" + ws_fymd3 + "',";
//				sQry = sQry + "ws_tymd1 = '" + ws_tymd1 + "',";
//				sQry = sQry + "ws_tymd2 = '" + ws_tymd2 + "',";
//				sQry = sQry + "ws_tymd3 = '" + ws_tymd3 + "',";
//				//UPGRADE_WARNING: ws_mamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_mamt1 = " + ws_mamt1 + ",";
//				//UPGRADE_WARNING: ws_mamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_mamt2 = " + ws_mamt2 + ",";
//				//UPGRADE_WARNING: ws_mamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_mamt3 = " + ws_mamt3 + ",";
//				//UPGRADE_WARNING: ws_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_gamt1 = " + ws_gamt1 + ",";
//				//UPGRADE_WARNING: ws_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ws_gamt2 = " + ws_gamt2 + ",";
//				sQry = sQry + "ws_gamt3 = " + ws_gamt3 + ",";

//				sQry = sQry + "dj_name1 = '" + dj_name1 + "',";
//				sQry = sQry + "dj_name2 = '" + dj_name2 + "',";
//				sQry = sQry + "dj_name3 = '" + dj_name3 + "',";
//				sQry = sQry + "dj_jumin1 = '" + dj_jumin1 + "',";
//				sQry = sQry + "dj_jumin2 = '" + dj_jumin2 + "',";
//				sQry = sQry + "dj_jumin3 = '" + dj_jumin3 + "',";
//				sQry = sQry + "dj_fymd1 = '" + dj_fymd1 + "',";
//				sQry = sQry + "dj_fymd2 = '" + dj_fymd2 + "',";
//				sQry = sQry + "dj_fymd3 = '" + dj_fymd3 + "',";
//				sQry = sQry + "dj_tymd1 = '" + dj_tymd1 + "',";
//				sQry = sQry + "dj_tymd2 = '" + dj_tymd2 + "',";
//				sQry = sQry + "dj_tymd3 = '" + dj_tymd3 + "',";
//				sQry = sQry + "dj_eja1 = '" + dj_eja1 + "',";
//				sQry = sQry + "dj_eja2 = '" + dj_eja2 + "',";
//				sQry = sQry + "dj_eja3 = '" + dj_eja3 + "',";
//				//UPGRADE_WARNING: dj_tamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_tamt1 = " + dj_tamt1 + ",";
//				//UPGRADE_WARNING: dj_tamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_tamt2 = " + dj_tamt2 + ",";
//				//UPGRADE_WARNING: dj_tamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_tamt3 = " + dj_tamt3 + ",";
//				//UPGRADE_WARNING: dj_wamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_wamt1 = " + dj_wamt1 + ",";
//				//UPGRADE_WARNING: dj_wamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_wamt2 = " + dj_wamt2 + ",";
//				sQry = sQry + "dj_wamt3 = " + dj_wamt3 + ",";
//				//UPGRADE_WARNING: dj_eamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_eamt1 = " + dj_eamt1 + ",";
//				//UPGRADE_WARNING: dj_eamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_eamt2 = " + dj_eamt2 + ",";
//				//UPGRADE_WARNING: dj_eamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_eamt3 = " + dj_eamt3 + ",";
//				//UPGRADE_WARNING: dj_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_gamt1 = " + dj_gamt1 + ",";
//				//UPGRADE_WARNING: dj_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "dj_gamt2 = " + dj_gamt2 + ",";
//				sQry = sQry + "dj_gamt3 = " + dj_gamt3 + ",";

//				sQry = sQry + "ld_name1 = '" + ld_name1 + "',";
//				sQry = sQry + "ld_name2 = '" + ld_name2 + "',";
//				sQry = sQry + "ld_name3 = '" + ld_name3 + "',";
//				sQry = sQry + "ld_jumin1 = '" + ld_jumin1 + "',";
//				sQry = sQry + "ld_jumin2 = '" + ld_jumin2 + "',";
//				sQry = sQry + "ld_jumin3 = '" + ld_jumin3 + "',";
//				sQry = sQry + "ld_hcode1 = '" + ld_hcode1 + "',";
//				sQry = sQry + "ld_hcode2 = '" + ld_hcode2 + "',";
//				sQry = sQry + "ld_hcode3 = '" + ld_hcode3 + "',";
//				//UPGRADE_WARNING: ld_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ld_hm1 = '" + ld_hm1 + "',";
//				//UPGRADE_WARNING: ld_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ld_hm2 = '" + ld_hm2 + "',";
//				sQry = sQry + "ld_hm3 = '" + ld_hm3 + "',";
//				sQry = sQry + "ld_addr1 = '" + ld_addr1 + "',";
//				sQry = sQry + "ld_addr2 = '" + ld_addr2 + "',";
//				sQry = sQry + "ld_addr3 = '" + ld_addr3 + "',";
//				sQry = sQry + "ld_fymd1 = '" + ld_fymd1 + "',";
//				sQry = sQry + "ld_fymd2 = '" + ld_fymd2 + "',";
//				sQry = sQry + "ld_fymd3 = '" + ld_fymd3 + "',";
//				sQry = sQry + "ld_tymd1 = '" + ld_tymd1 + "',";
//				sQry = sQry + "ld_tymd2 = '" + ld_tymd2 + "',";
//				sQry = sQry + "ld_tymd3 = '" + ld_tymd3 + "',";
//				//UPGRADE_WARNING: ld_bamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ld_bamt1 = " + ld_bamt1 + ",";
//				//UPGRADE_WARNING: ld_bamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ld_bamt2 = " + ld_bamt2 + ",";
//				sQry = sQry + "ld_bamt3 = " + ld_bamt3 + "";

//				sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";

//				oRecordSet.DoQuery(sQry);

//			} else {

//				////신규
//				sQry = "INSERT INTO [p_seoyhouse]";
//				sQry = sQry + " (";
//				sQry = sQry + "saup,";
//				sQry = sQry + "yyyy,";
//				sQry = sQry + "sabun,";
//				sQry = sQry + "ws_name1, ";
//				sQry = sQry + "ws_name2, ";
//				sQry = sQry + "ws_name3, ";
//				sQry = sQry + "ws_jumin1, ";
//				sQry = sQry + "ws_jumin2, ";
//				sQry = sQry + "ws_jumin3, ";
//				sQry = sQry + "ws_hcode1, ";
//				sQry = sQry + "ws_hcode2, ";
//				sQry = sQry + "ws_hcode3, ";
//				sQry = sQry + "ws_hm1, ";
//				sQry = sQry + "ws_hm2, ";
//				sQry = sQry + "ws_hm3, ";
//				sQry = sQry + "ws_addr1, ";
//				sQry = sQry + "ws_addr2, ";
//				sQry = sQry + "ws_addr3, ";
//				sQry = sQry + "ws_fymd1, ";
//				sQry = sQry + "ws_fymd2, ";
//				sQry = sQry + "ws_fymd3, ";
//				sQry = sQry + "ws_tymd1, ";
//				sQry = sQry + "ws_tymd2, ";
//				sQry = sQry + "ws_tymd3, ";
//				sQry = sQry + "ws_mamt1, ";
//				sQry = sQry + "ws_mamt2, ";
//				sQry = sQry + "ws_mamt3, ";
//				sQry = sQry + "ws_gamt1, ";
//				sQry = sQry + "ws_gamt2, ";
//				sQry = sQry + "ws_gamt3, ";

//				sQry = sQry + "dj_name1, ";
//				sQry = sQry + "dj_name2, ";
//				sQry = sQry + "dj_name3, ";
//				sQry = sQry + "dj_jumin1, ";
//				sQry = sQry + "dj_jumin2, ";
//				sQry = sQry + "dj_jumin3, ";
//				sQry = sQry + "dj_fymd1, ";
//				sQry = sQry + "dj_fymd2, ";
//				sQry = sQry + "dj_fymd3, ";
//				sQry = sQry + "dj_tymd1, ";
//				sQry = sQry + "dj_tymd2, ";
//				sQry = sQry + "dj_tymd3, ";
//				sQry = sQry + "dj_eja1, ";
//				sQry = sQry + "dj_eja2, ";
//				sQry = sQry + "dj_eja3, ";
//				sQry = sQry + "dj_tamt1, ";
//				sQry = sQry + "dj_tamt2, ";
//				sQry = sQry + "dj_tamt3, ";
//				sQry = sQry + "dj_wamt1, ";
//				sQry = sQry + "dj_wamt2, ";
//				sQry = sQry + "dj_wamt3, ";
//				sQry = sQry + "dj_eamt1, ";
//				sQry = sQry + "dj_eamt2, ";
//				sQry = sQry + "dj_eamt3, ";
//				sQry = sQry + "dj_gamt1, ";
//				sQry = sQry + "dj_gamt2, ";
//				sQry = sQry + "dj_gamt3, ";

//				sQry = sQry + "ld_name1, ";
//				sQry = sQry + "ld_name2, ";
//				sQry = sQry + "ld_name3, ";
//				sQry = sQry + "ld_jumin1, ";
//				sQry = sQry + "ld_jumin2, ";
//				sQry = sQry + "ld_jumin3, ";
//				sQry = sQry + "ld_hcode1, ";
//				sQry = sQry + "ld_hcode2, ";
//				sQry = sQry + "ld_hcode3, ";
//				sQry = sQry + "ld_hm1, ";
//				sQry = sQry + "ld_hm2, ";
//				sQry = sQry + "ld_hm3, ";
//				sQry = sQry + "ld_addr1, ";
//				sQry = sQry + "ld_addr2, ";
//				sQry = sQry + "ld_addr3, ";
//				sQry = sQry + "ld_fymd1, ";
//				sQry = sQry + "ld_fymd2, ";
//				sQry = sQry + "ld_fymd3, ";
//				sQry = sQry + "ld_tymd1, ";
//				sQry = sQry + "ld_tymd2, ";
//				sQry = sQry + "ld_tymd3, ";
//				sQry = sQry + "ld_bamt1, ";
//				sQry = sQry + "ld_bamt2, ";
//				sQry = sQry + "ld_bamt3 ";
//				sQry = sQry + " ) ";
//				sQry = sQry + "VALUES(";

//				sQry = sQry + "'" + saup + "',";
//				sQry = sQry + "'" + yyyy + "',";
//				sQry = sQry + "'" + sabun + "',";
//				sQry = sQry + "'" + ws_name1 + "',";
//				sQry = sQry + "'" + ws_name2 + "',";
//				sQry = sQry + "'" + ws_name3 + "',";
//				sQry = sQry + "'" + ws_jumin1 + "',";
//				sQry = sQry + "'" + ws_jumin2 + "',";
//				sQry = sQry + "'" + ws_jumin3 + "',";
//				sQry = sQry + "'" + ws_hcode1 + "',";
//				sQry = sQry + "'" + ws_hcode2 + "',";
//				sQry = sQry + "'" + ws_hcode3 + "',";
//				//UPGRADE_WARNING: ws_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_hm1 + ",";
//				//UPGRADE_WARNING: ws_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_hm2 + ",";
//				sQry = sQry + ws_hm3 + ",";
//				sQry = sQry + "'" + ws_addr1 + "',";
//				sQry = sQry + "'" + ws_addr2 + "',";
//				sQry = sQry + "'" + ws_addr3 + "',";
//				sQry = sQry + "'" + ws_fymd1 + "',";
//				sQry = sQry + "'" + ws_fymd2 + "',";
//				sQry = sQry + "'" + ws_fymd3 + "',";
//				sQry = sQry + "'" + ws_tymd1 + "',";
//				sQry = sQry + "'" + ws_tymd2 + "',";
//				sQry = sQry + "'" + ws_tymd3 + "',";
//				//UPGRADE_WARNING: ws_mamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_mamt1 + ",";
//				//UPGRADE_WARNING: ws_mamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_mamt2 + ",";
//				//UPGRADE_WARNING: ws_mamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_mamt3 + ",";
//				//UPGRADE_WARNING: ws_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_gamt1 + ",";
//				//UPGRADE_WARNING: ws_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ws_gamt2 + ",";
//				sQry = sQry + ws_gamt3 + ",";

//				sQry = sQry + "'" + dj_name1 + "',";
//				sQry = sQry + "'" + dj_name2 + "',";
//				sQry = sQry + "'" + dj_name3 + "',";
//				sQry = sQry + "'" + dj_jumin1 + "',";
//				sQry = sQry + "'" + dj_jumin2 + "',";
//				sQry = sQry + "'" + dj_jumin3 + "',";
//				sQry = sQry + "'" + dj_fymd1 + "',";
//				sQry = sQry + "'" + dj_fymd2 + "',";
//				sQry = sQry + "'" + dj_fymd3 + "',";
//				sQry = sQry + "'" + dj_tymd1 + "',";
//				sQry = sQry + "'" + dj_tymd2 + "',";
//				sQry = sQry + "'" + dj_tymd3 + "',";
//				sQry = sQry + "'" + dj_eja1 + "',";
//				sQry = sQry + "'" + dj_eja2 + "',";
//				sQry = sQry + "'" + dj_eja3 + "',";
//				//UPGRADE_WARNING: dj_tamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_tamt1 + ",";
//				//UPGRADE_WARNING: dj_tamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_tamt2 + ",";
//				//UPGRADE_WARNING: dj_tamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_tamt3 + ",";
//				//UPGRADE_WARNING: dj_wamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_wamt1 + ",";
//				//UPGRADE_WARNING: dj_wamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_wamt2 + ",";
//				sQry = sQry + dj_wamt3 + ",";
//				//UPGRADE_WARNING: dj_eamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_eamt1 + ",";
//				//UPGRADE_WARNING: dj_eamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_eamt2 + ",";
//				//UPGRADE_WARNING: dj_eamt3 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_eamt3 + ",";
//				//UPGRADE_WARNING: dj_gamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_gamt1 + ",";
//				//UPGRADE_WARNING: dj_gamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + dj_gamt2 + ",";
//				sQry = sQry + dj_gamt3 + ",";

//				sQry = sQry + "'" + ld_name1 + "',";
//				sQry = sQry + "'" + ld_name2 + "',";
//				sQry = sQry + "'" + ld_name3 + "',";
//				sQry = sQry + "'" + ld_jumin1 + "',";
//				sQry = sQry + "'" + ld_jumin2 + "',";
//				sQry = sQry + "'" + ld_jumin3 + "',";
//				sQry = sQry + "'" + ld_hcode1 + "',";
//				sQry = sQry + "'" + ld_hcode2 + "',";
//				sQry = sQry + "'" + ld_hcode3 + "',";
//				//UPGRADE_WARNING: ld_hm1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ld_hm1 + ",";
//				//UPGRADE_WARNING: ld_hm2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ld_hm2 + ",";
//				sQry = sQry + ld_hm3 + ",";
//				sQry = sQry + "'" + ld_addr1 + "',";
//				sQry = sQry + "'" + ld_addr2 + "',";
//				sQry = sQry + "'" + ld_addr3 + "',";
//				sQry = sQry + "'" + ld_fymd1 + "',";
//				sQry = sQry + "'" + ld_fymd2 + "',";
//				sQry = sQry + "'" + ld_fymd3 + "',";
//				sQry = sQry + "'" + ld_tymd1 + "',";
//				sQry = sQry + "'" + ld_tymd2 + "',";
//				sQry = sQry + "'" + ld_tymd3 + "',";
//				//UPGRADE_WARNING: ld_bamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ld_bamt1 + ",";
//				//UPGRADE_WARNING: ld_bamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ld_bamt2 + ",";
//				sQry = sQry + ld_bamt3;
//				sQry = sQry + " ) ";

//				oRecordSet.DoQuery(sQry);
//			}


//			PH_PY413_FormItemEnabled();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			PH_PY413_MTX01();

//			return;
//			PH_PY413_SAVE_Exit:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY413_SAVE_Error:
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY413_Delete()
//		{
//			////선택된 자료 삭제

//			string sabun = null;
//			string saup = null;
//			string yyyy = null;
//			string kname = null;

//			short i = 0;
//			short cnt = 0;

//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			kname = oForm.Items.Item("FullName").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sabun = Strings.Trim(Conversion.Str(oForm.Items.Item("MSTCOD").Specific.VALUE));

//			sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			cnt = oRecordSet.Fields.Item(0).Value;
//			if (cnt > 0) {

//				if (string.IsNullOrEmpty(Strings.Trim(yyyy))) {
//					MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY413_Delete_Exit;
//				}

//				if (string.IsNullOrEmpty(Strings.Trim(saup))) {
//					MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY413_Delete_Exit;
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(sabun))) {
//					MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY413_Delete_Exit;
//				}


//				if (MDC_Globals.Sbo_Application.MessageBox(" 선택한사원('" + kname + "')의 모든자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1")) {
//					sQry = "Delete From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//					oRecordSet.DoQuery(sQry);
//				}
//			}

//			oForm.Freeze(false);

//			PH_PY413_MTX01();

//			oForm.ActiveItem = "MSTCOD";

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;


//			return;
//			PH_PY413_Delete_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			return;
//			PH_PY413_Delete_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY413_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY413_TitleSetting(ref int iRow)
//		{
//			//    Dim i               As Long
//			//    Dim j               As Long
//			//    Dim sQry            As String
//			//
//			//    Dim COLNAM(5)       As String
//			//
//			//    Dim oColumn         As SAPbouiCOM.EditTextColumn
//			//    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
//			//
//			//    Dim oRecordSet  As SAPbobsCOM.Recordset
//			//
//			//    On Error GoTo Error_Message
//			//
//			//    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
//			//
//			//    oForm.Freeze True
//			//
//			//    COLNAM(0) = "년도"
//			//    COLNAM(1) = "부서"
//			//    COLNAM(2) = "담당"
//			//    COLNAM(3) = "사번"
//			//    COLNAM(4) = "성명"
//			//    COLNAM(5) = "직급"
//			//
//			//    For i = 0 To UBound(COLNAM)
//			//        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
//			//        oGrid1.Columns.Item(i).Editable = False
//			//
//			//        oGrid1.Columns.Item(i).RightJustified = True
//			//
//			//    Next i
//			//
//			//    oGrid1.AutoResizeColumns
//			//
//			//    oForm.Freeze False
//			//
//			//    Set oColumn = Nothing
//			//
//			//    Exit Sub
//			//
//			//Error_Message:
//			//    oForm.Freeze False
//			//    Set oColumn = Nothing
//			//    Sbo_Application.SetStatusBarMessage "PH_PY413_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
//		}
//	}
//}
