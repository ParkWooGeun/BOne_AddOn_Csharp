﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전근무지등록
    /// </summary>
    internal class PH_PY401 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY401;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY401.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY401_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY401");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY401_CreateItems();
                PH_PY401_FormItemEnabled();
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
        private void PH_PY401_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY401");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY401");
                oDS_PH_PY401 = oForm.DataSources.DataTables.Item("PH_PY401");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("직급", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                
                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
                oForm.DataSources.UserDataSources.Item("Year").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
               
                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 사업자등록번호1,2
                oForm.DataSources.UserDataSources.Add("entno1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("entno1").Specific.DataBind.SetBound(true, "", "entno1");
                oForm.DataSources.UserDataSources.Add("entno2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("entno2").Specific.DataBind.SetBound(true, "", "entno2");

                // 근무처명1,2
                oForm.DataSources.UserDataSources.Add("servcomp1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("servcomp1").Specific.DataBind.SetBound(true, "", "servcomp1");
                oForm.DataSources.UserDataSources.Add("servcomp2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("servcomp2").Specific.DataBind.SetBound(true, "", "servcomp2");

                // 시작근무일1,2
                oForm.DataSources.UserDataSources.Add("symd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("symd1").Specific.DataBind.SetBound(true, "", "symd1");
                oForm.DataSources.UserDataSources.Add("symd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("symd2").Specific.DataBind.SetBound(true, "", "symd2");

                // 종료근무일1,2
                oForm.DataSources.UserDataSources.Add("eymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("eymd1").Specific.DataBind.SetBound(true, "", "eymd1");
                oForm.DataSources.UserDataSources.Add("eymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("eymd2").Specific.DataBind.SetBound(true, "", "eymd2");

                // 급여총액1,2
                oForm.DataSources.UserDataSources.Add("payrtot1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("payrtot1").Specific.DataBind.SetBound(true, "", "payrtot1");
                oForm.DataSources.UserDataSources.Add("payrtot2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("payrtot2").Specific.DataBind.SetBound(true, "", "payrtot2");

                // 상여총액1,2
                oForm.DataSources.UserDataSources.Add("bnstot1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bnstot1").Specific.DataBind.SetBound(true, "", "bnstot1");
                oForm.DataSources.UserDataSources.Add("bnstot2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bnstot2").Specific.DataBind.SetBound(true, "", "bnstot2");

                // 국외근로1,2
                oForm.DataSources.UserDataSources.Add("fwork1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fwork1").Specific.DataBind.SetBound(true, "", "fwork1");
                oForm.DataSources.UserDataSources.Add("fwork2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fwork2").Specific.DataBind.SetBound(true, "", "fwork2");

                // 야간근로수당등1,2
                oForm.DataSources.UserDataSources.Add("ndtalw1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ndtalw1").Specific.DataBind.SetBound(true, "", "ndtalw1");
                oForm.DataSources.UserDataSources.Add("ndtalw2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ndtalw2").Specific.DataBind.SetBound(true, "", "ndtalw2");

                // 기타비과세1,2
                oForm.DataSources.UserDataSources.Add("etcntax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("etcntax1").Specific.DataBind.SetBound(true, "", "etcntax1");
                oForm.DataSources.UserDataSources.Add("etcntax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("etcntax2").Specific.DataBind.SetBound(true, "", "etcntax2");

                // 중식수당1,2
                oForm.DataSources.UserDataSources.Add("lnchalw1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("lnchalw1").Specific.DataBind.SetBound(true, "", "lnchalw1");
                oForm.DataSources.UserDataSources.Add("lnchalw2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("lnchalw2").Specific.DataBind.SetBound(true, "", "lnchalw2");

                // 외국납부세액1,2
                oForm.DataSources.UserDataSources.Add("ftaxamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ftaxamt1").Specific.DataBind.SetBound(true, "", "ftaxamt1");
                oForm.DataSources.UserDataSources.Add("ftaxamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ftaxamt2").Specific.DataBind.SetBound(true, "", "ftaxamt2");

                // 저축세액공제1,2
                oForm.DataSources.UserDataSources.Add("savtaxddc1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("savtaxddc1").Specific.DataBind.SetBound(true, "", "savtaxddc1");
                oForm.DataSources.UserDataSources.Add("savtaxddc2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("savtaxddc2").Specific.DataBind.SetBound(true, "", "savtaxddc2");

                // 소득세1,2
                oForm.DataSources.UserDataSources.Add("incmtax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("incmtax1").Specific.DataBind.SetBound(true, "", "incmtax1");
                oForm.DataSources.UserDataSources.Add("indmtax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("indmtax2").Specific.DataBind.SetBound(true, "", "indmtax2");

                // 농어촌특별세1,2
                oForm.DataSources.UserDataSources.Add("fvsptax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fvsptax1").Specific.DataBind.SetBound(true, "", "fvsptax1");
                oForm.DataSources.UserDataSources.Add("fvsptax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fvsptax2").Specific.DataBind.SetBound(true, "", "fvsptax2");

                // 주민세1,2
                oForm.DataSources.UserDataSources.Add("residtax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("residtax1").Specific.DataBind.SetBound(true, "", "residtax1");
                oForm.DataSources.UserDataSources.Add("residtax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("residtax2").Specific.DataBind.SetBound(true, "", "residtax2");

                // 건강보험1,2
                oForm.DataSources.UserDataSources.Add("medcinsr1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcinsr1").Specific.DataBind.SetBound(true, "", "medcinsr1");
                oForm.DataSources.UserDataSources.Add("medcinsr2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcinsr2").Specific.DataBind.SetBound(true, "", "medcinsr2");

                // 고용보험1,2
                oForm.DataSources.UserDataSources.Add("asopinsr1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("asopinsr1").Specific.DataBind.SetBound(true, "", "asopinsr1");
                oForm.DataSources.UserDataSources.Add("asopinsr2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("asopinsr2").Specific.DataBind.SetBound(true, "", "asopinsr2");

                // 연금보험1,2
                oForm.DataSources.UserDataSources.Add("annuboamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("annuboamt1").Specific.DataBind.SetBound(true, "", "annuboamt1");
                oForm.DataSources.UserDataSources.Add("annuboamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("annuboamt2").Specific.DataBind.SetBound(true, "", "annuboamt2");

                // 인정상여1,2
                oForm.DataSources.UserDataSources.Add("injungpay1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("injungpay1").Specific.DataBind.SetBound(true, "", "injungpay1");
                oForm.DataSources.UserDataSources.Add("injungpay2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("injungpay2").Specific.DataBind.SetBound(true, "", "injungpay2");

                // 임원퇴직소득금액 한도초과액1,2
                oForm.DataSources.UserDataSources.Add("overpay1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("overpay1").Specific.DataBind.SetBound(true, "", "overpay1");
                oForm.DataSources.UserDataSources.Add("overpay2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("overpay2").Specific.DataBind.SetBound(true, "", "overpay2");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY401_FormItemEnabled()
        {
            try
            {
                oForm.Items.Item("MSTCOD").Specific.Value = "";
                oForm.Items.Item("FullName").Specific.Value = "";
                oForm.Items.Item("TeamName").Specific.Value = "";
                oForm.Items.Item("RspName").Specific.Value = "";
                oForm.Items.Item("ClsName").Specific.Value = "";

                oForm.DataSources.UserDataSources.Item("entno1").Value = "";
                oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
                oForm.DataSources.UserDataSources.Item("symd1").Value = "";
                oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

                oForm.DataSources.UserDataSources.Item("payrtot1").Value = "0";
                oForm.DataSources.UserDataSources.Item("bnstot1").Value = "0";
                oForm.DataSources.UserDataSources.Item("fwork1").Value = "0";
                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = "0";
                oForm.DataSources.UserDataSources.Item("etcntax1").Value = "0";
                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = "0";
                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = "0";
                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = "0";
                oForm.DataSources.UserDataSources.Item("incmtax1").Value = "0";
                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = "0";
                oForm.DataSources.UserDataSources.Item("residtax1").Value = "0";
                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = "0";
                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = "0";
                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = "0";
                oForm.DataSources.UserDataSources.Item("injungpay1").Value = "0";
                oForm.DataSources.UserDataSources.Item("overpay1").Value = "0";

                oForm.DataSources.UserDataSources.Item("entno2").Value = "";
                oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
                oForm.DataSources.UserDataSources.Item("symd2").Value = "";
                oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

                oForm.DataSources.UserDataSources.Item("payrtot2").Value = "0";
                oForm.DataSources.UserDataSources.Item("bnstot2").Value = "0";
                oForm.DataSources.UserDataSources.Item("fwork2").Value = "0";
                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = "0";
                oForm.DataSources.UserDataSources.Item("etcntax2").Value = "0";
                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = "0";
                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = "0";
                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = "0";
                oForm.DataSources.UserDataSources.Item("indmtax2").Value = "0";
                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = "0";
                oForm.DataSources.UserDataSources.Item("residtax2").Value = "0";
                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = "0";
                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = "0";
                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = "0";
                oForm.DataSources.UserDataSources.Item("injungpay2").Value = "0";
                oForm.DataSources.UserDataSources.Item("overpay2").Value = "0";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY401_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY401_DataFind
        /// </summary>
        private void PH_PY401_DataFind()
        {
            string sQry;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    errMessage = "년도가 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                PH_PY401_FormItemEnabled();

                sQry = "Exec PH_PY401_01 '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "','" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "'";
                oDS_PH_PY401.ExecuteQuery(sQry);

                PH_PY401_TitleSetting();

                oForm.EnableMenu("1282", true);     //문서추가 활성
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY401_SAVE
        /// </summary>
        private void PH_PY401_SAVE()
        {
            int ErrNum = 0;
            string sQry;
            string CLTCOD;
            string MSTCOD;
            string FullName;
            string YEAR;
            string entno1;
            string servcomp1;
            string symd1;
            string eymd1;
            string entno2;
            string servcomp2;
            string symd2;
            string eymd2;
            double payrtot1;
            double bnstot1;
            double fwork1;
            double ndtalw1;
            double etcntax1;
            double lnchalw1;
            double ftaxamt1;
            double savtaxddc1;
            double incmtax1;
            double fvsptax1;
            double residtax1;
            double medcinsr1;
            double asopinsr1;
            double annuboamt1;
            double injungpay1;
            double overpay1;
            double payrtot2;
            double bnstot2;
            double fwork2;
            double ndtalw2;
            double etcntax2;
            double lnchalw2;
            double ftaxamt2;
            double savtaxddc2;
            double indmtax2;
            double fvsptax2;
            double residtax2;
            double medcinsr2;
            double asopinsr2;
            double annuboamt2;
            double injungpay2;
            double overpay2;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                entno1 = oForm.Items.Item("entno1").Specific.Value.ToString().Trim();
                servcomp1 = oForm.Items.Item("servcomp1").Specific.Value.ToString().Trim();
                symd1 = oForm.Items.Item("symd1").Specific.Value.ToString().Trim();
                eymd1 = oForm.Items.Item("eymd1").Specific.Value.ToString().Trim();
                payrtot1 = Convert.ToDouble(oForm.Items.Item("payrtot1").Specific.Value.ToString().Trim());
                bnstot1 = Convert.ToDouble(oForm.Items.Item("bnstot1").Specific.Value.ToString().Trim());
                fwork1 = Convert.ToDouble(oForm.Items.Item("fwork1").Specific.Value.ToString().Trim());
                ndtalw1 = Convert.ToDouble(oForm.Items.Item("ndtalw1").Specific.Value.ToString().Trim());
                etcntax1 = Convert.ToDouble(oForm.Items.Item("etcntax1").Specific.Value.ToString().Trim());
                lnchalw1 = Convert.ToDouble(oForm.Items.Item("lnchalw1").Specific.Value.ToString().Trim());
                ftaxamt1 = Convert.ToDouble(oForm.Items.Item("ftaxamt1").Specific.Value.ToString().Trim());
                savtaxddc1 = Convert.ToDouble(oForm.Items.Item("savtaxddc1").Specific.Value.ToString().Trim());
                incmtax1 = Convert.ToDouble(oForm.Items.Item("incmtax1").Specific.Value.ToString().Trim());
                fvsptax1 = Convert.ToDouble(oForm.Items.Item("fvsptax1").Specific.Value.ToString().Trim());
                residtax1 = Convert.ToDouble(oForm.Items.Item("residtax1").Specific.Value.ToString().Trim());
                medcinsr1 = Convert.ToDouble(oForm.Items.Item("medcinsr1").Specific.Value.ToString().Trim());
                asopinsr1 = Convert.ToDouble(oForm.Items.Item("asopinsr1").Specific.Value.ToString().Trim());
                annuboamt1 = Convert.ToDouble(oForm.Items.Item("annuboamt1").Specific.Value.ToString().Trim());
                injungpay1 = Convert.ToDouble(oForm.Items.Item("injungpay1").Specific.Value.ToString().Trim());
                overpay1 = Convert.ToDouble(oForm.Items.Item("overpay1").Specific.Value.ToString().Trim());

                entno2 = oForm.Items.Item("entno2").Specific.Value.ToString().Trim();
                servcomp2 = oForm.Items.Item("servcomp2").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(oForm.Items.Item("symd2").Specific.Value.ToString().Trim()))
                {
                    symd2 = "";
                    eymd2 = "";
                }
                else
                {
                    symd2 = oForm.Items.Item("symd2").Specific.Value.ToString().Trim();
                    eymd2 = oForm.Items.Item("eymd2").Specific.Value.ToString().Trim();
                }

                payrtot2 = Convert.ToDouble(oForm.Items.Item("payrtot2").Specific.Value.ToString().Trim());
                bnstot2 = Convert.ToDouble(oForm.Items.Item("bnstot2").Specific.Value.ToString().Trim());
                fwork2 = Convert.ToDouble(oForm.Items.Item("fwork2").Specific.Value.ToString().Trim());
                ndtalw2 = Convert.ToDouble(oForm.Items.Item("ndtalw2").Specific.Value.ToString().Trim());
                etcntax2 = Convert.ToDouble(oForm.Items.Item("etcntax2").Specific.Value.ToString().Trim());
                lnchalw2 = Convert.ToDouble(oForm.Items.Item("lnchalw2").Specific.Value.ToString().Trim());
                ftaxamt2 = Convert.ToDouble(oForm.Items.Item("ftaxamt2").Specific.Value.ToString().Trim());
                savtaxddc2 = Convert.ToDouble(oForm.Items.Item("savtaxddc2").Specific.Value.ToString().Trim());
                indmtax2 = Convert.ToDouble(oForm.Items.Item("indmtax2").Specific.Value.ToString().Trim());
                fvsptax2 = Convert.ToDouble(oForm.Items.Item("fvsptax2").Specific.Value.ToString().Trim());
                residtax2 = Convert.ToDouble(oForm.Items.Item("residtax2").Specific.Value.ToString().Trim());
                medcinsr2 = Convert.ToDouble(oForm.Items.Item("medcinsr2").Specific.Value.ToString().Trim());
                asopinsr2 = Convert.ToDouble(oForm.Items.Item("asopinsr2").Specific.Value.ToString().Trim());
                annuboamt2 = Convert.ToDouble(oForm.Items.Item("annuboamt2").Specific.Value.ToString().Trim());
                injungpay2 = Convert.ToDouble(oForm.Items.Item("injungpay2").Specific.Value.ToString().Trim());
                overpay2 = Convert.ToDouble(oForm.Items.Item("overpay2").Specific.Value.ToString().Trim());

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YEAR = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();


                if (string.IsNullOrWhiteSpace(CLTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(YEAR))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(MSTCOD))
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    //신규
                    sQry = "INSERT INTO [p_sbservcomp]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "entno1,";
                    sQry += "servcomp1,";
                    sQry += "symd1,";
                    sQry += "eymd1,";
                    sQry += "payrtot1,";
                    sQry += "bnstot1,";
                    sQry += "fwork1,";
                    sQry += "ndtalw1,";
                    sQry += "etcntax1,";
                    sQry += "lnchalw1,";
                    sQry += "ftaxamt1,";
                    sQry += "savtaxddc1,";
                    sQry += "incmtax1,";
                    sQry += "fvsptax1,";
                    sQry += "residtax1,";
                    sQry += "medcinsr1,";
                    sQry += "asopinsr1,";
                    sQry += "annuboamt1,";
                    sQry += "entno2,";
                    sQry += "servcomp2,";
                    sQry += "symd2,";
                    sQry += "eymd2,";
                    sQry += "payrtot2,";
                    sQry += "bnstot2,";
                    sQry += "fwork2,";
                    sQry += "ndtalw2,";
                    sQry += "etcntax2,";
                    sQry += "lnchalw2,";
                    sQry += "ftaxamt2,";
                    sQry += "savtaxddc2,";
                    sQry += "indmtax2,";
                    sQry += "fvsptax2,";
                    sQry += "residtax2,";
                    sQry += "medcinsr2,";
                    sQry += "asopinsr2,";
                    sQry += "annuboamt2,";
                    sQry += "jscntddc1,";
                    sQry += "jscntddc2,";
                    sQry += "injungpay1,";
                    sQry += "overpay1,"; 
                    sQry += "injungpay2,";
                    sQry += "overpay2";
                    sQry += " ) ";
                    sQry += "VALUES(";
                    sQry += "'" + CLTCOD + "',";
                    sQry += "'" + YEAR + "',";
                    sQry += "'" + MSTCOD + "',";
                    sQry += "'" + entno1 + "',";
                    sQry += "'" + servcomp1 + "',";
                    sQry += "'" + symd1 + "',";
                    sQry += "'" + eymd1 + "',";
                    sQry += payrtot1 + ",";
                    sQry += bnstot1 + ",";
                    sQry += fwork1 + ",";
                    sQry += ndtalw1 + ",";
                    sQry += etcntax1 + ",";
                    sQry += lnchalw1 + ",";
                    sQry += ftaxamt1 + ",";
                    sQry += savtaxddc1 + ",";
                    sQry += incmtax1 + ",";
                    sQry += fvsptax1 + ",";
                    sQry += residtax1 + ",";
                    sQry += medcinsr1 + ",";
                    sQry += asopinsr1 + ",";
                    sQry += annuboamt1 + ",";
                    sQry += "'" + entno2 + "',";
                    sQry += "'" + servcomp2 + "',";
                    sQry += "'" + symd2 + "',";
                    sQry += "'" + eymd2 + "',";
                    sQry += payrtot2 + ",";
                    sQry += bnstot2 + ",";
                    sQry += fwork2 + ",";
                    sQry += ndtalw2 + ",";
                    sQry += etcntax2 + ",";
                    sQry += lnchalw2 + ",";
                    sQry += ftaxamt2 + ",";
                    sQry += savtaxddc2 + ",";
                    sQry += indmtax2 + ",";
                    sQry += fvsptax2 + ",";
                    sQry += residtax2 + ",";
                    sQry += medcinsr2 + ",";
                    sQry += asopinsr2 + ",";
                    sQry += annuboamt2 + ",";
                    sQry += "0, 0,";
                    sQry += injungpay1 + ",";
                    sQry += overpay1 + ",";
                    sQry += injungpay2 + ",";
                    sQry += overpay2;
                    sQry += " ) ";
                    oRecordSet.DoQuery(sQry);

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY401_DataFind();
                }
                else
                {
                    //수정
                    sQry = "Update [p_sbservcomp] set ";
                    sQry += "entno1 = '" + entno1 + "',";
                    sQry += "servcomp1 = '" + servcomp1 + "',";
                    sQry += "symd1 = '" + symd1 + "',";
                    sQry += "eymd1 = '" + eymd1 + "',";
                    sQry += "payrtot1 = " + payrtot1 + ",";
                    sQry += "bnstot1 = " + bnstot1 + ",";
                    sQry += "fwork1 = " + fwork1 + ",";
                    sQry += "ndtalw1 = " + ndtalw1 + ",";
                    sQry += "etcntax1 = " + etcntax1 + ",";
                    sQry += "lnchalw1 = " + lnchalw1 + ",";
                    sQry += "ftaxamt1 = " + ftaxamt1 + ",";
                    sQry += "savtaxddc1 = " + savtaxddc1 + ",";
                    sQry += "incmtax1 = " + incmtax1 + ",";
                    sQry += "fvsptax1 = " + fvsptax1 + ",";
                    sQry += "residtax1 = " + residtax1 + ",";
                    sQry += "medcinsr1 = " + medcinsr1 + ",";
                    sQry += "asopinsr1 = " + asopinsr1 + ",";
                    sQry += "annuboamt1 =" + annuboamt1 + ",";
                    sQry += "entno2 = '" + entno2 + "',";
                    sQry += "servcomp2 = '" + servcomp2 + "',";
                    sQry += "symd2 = '" + symd2 + "',";
                    sQry += "eymd2 = '" + eymd2 + "',";
                    sQry += "payrtot2 = " + payrtot2 + ",";
                    sQry += "bnstot2= " + bnstot2 + ",";
                    sQry += "fwork2 = " + fwork2 + ",";
                    sQry += "ndtalw2 = " + ndtalw2 + ",";
                    sQry += "etcntax2 = " + etcntax2 + ",";
                    sQry += "lnchalw2 = " + lnchalw2 + ",";
                    sQry += "ftaxamt2 = " + ftaxamt2 + ",";
                    sQry += "savtaxddc2 = " + savtaxddc2 + ",";
                    sQry += "indmtax2 = " + indmtax2 + ",";
                    sQry += "fvsptax2 = " + fvsptax2 + ",";
                    sQry += "residtax2 = " + residtax2 + ",";
                    sQry += "medcinsr2 = " + medcinsr2 + ",";
                    sQry += "asopinsr2 = " + asopinsr2 + ",";
                    sQry += "annuboamt2 = " + annuboamt2 + ",";
                    sQry += "injungpay1 = " + injungpay1 + ",";
                    sQry += "overpay1 = " + overpay1 + ",";
                    sQry += "injungpay2 = " + injungpay2 + ",";
                    sQry += "overpay2 = " + overpay2;
                    sQry += " Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY401_DataFind();
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장코드를 확인 하세요.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("년도를 확인 하세요.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("사원코드를 확인 하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY401_Delete 데이타 삭제
        /// </summary>
        private void PH_PY401_Delete()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    if (oDS_PH_PY401.Rows.Count > 0)
                    {
                        sQry = " Delete From [p_sbservcomp] Where  saup = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "' And yyyy = '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "' And sabun = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY401_DataFind();
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY401_TitleSetting
        /// </summary>
        private void PH_PY401_TitleSetting()
        {
            int i;
            string[] COLNAM = new string[6];

            try
            {
                //그리드 콤보박스
                COLNAM[0] = "년도";
                COLNAM[1] = "부서";
                COLNAM[2] = "담당";
                COLNAM[3] = "사번";
                COLNAM[4] = "성명";
                COLNAM[5] = "직급";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 && i < COLNAM.Length)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
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
            string sQry;
            string Result;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY401_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY401_SAVE();
                        }
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY401_Delete();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        if (pVal.ItemUID == "Grid01")
                        {
                        }
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                string sQry;
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);
                            sQry = "EXEC PH_PY401_02 '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "', '" + oDS_PH_PY401.Columns.Item("Year").Cells.Item(pVal.Row).Value;
                            sQry += "', '" + oDS_PH_PY401.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value + "'";
                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.RecordCount == 0)
                            {
                                oForm.Items.Item("MSTCOD").Specific.Value = oDS_PH_PY401.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                                oForm.Items.Item("FullName").Specific.Value = oDS_PH_PY401.Columns.Item("FullName").Cells.Item(pVal.Row).Value;

                                oForm.DataSources.UserDataSources.Item("entno1").Value = "";
                                oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
                                oForm.DataSources.UserDataSources.Item("symd1").Value = "";
                                oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

                                oForm.DataSources.UserDataSources.Item("payrtot1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("bnstot1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("fwork1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("etcntax1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("incmtax1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("residtax1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("injungpay1").Value = "0";
                                oForm.DataSources.UserDataSources.Item("overpay1").Value = "0";

                                oForm.DataSources.UserDataSources.Item("entno2").Value = "";
                                oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
                                oForm.DataSources.UserDataSources.Item("symd2").Value = "";
                                oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

                                oForm.DataSources.UserDataSources.Item("payrtot2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("bnstot2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("fwork2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("etcntax2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("indmtax2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("residtax2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("injungpay2").Value = "0";
                                oForm.DataSources.UserDataSources.Item("overpay2").Value = "0";

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {
                                oForm.Items.Item("MSTCOD").Specific.Value = oRecordSet.Fields.Item("MSTCOD").Value;
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;

                                oForm.DataSources.UserDataSources.Item("entno1").Value = oRecordSet.Fields.Item("entno1").Value;
                                oForm.DataSources.UserDataSources.Item("servcomp1").Value = oRecordSet.Fields.Item("servcomp1").Value;
                                oForm.DataSources.UserDataSources.Item("symd1").Value = Convert.ToDateTime(oRecordSet.Fields.Item("symd1").Value.ToString().Trim()).ToString("yyyyMMdd");
                                oForm.DataSources.UserDataSources.Item("eymd1").Value = Convert.ToDateTime(oRecordSet.Fields.Item("eymd1").Value.ToString().Trim()).ToString("yyyyMMdd");

                                oForm.DataSources.UserDataSources.Item("payrtot1").Value = oRecordSet.Fields.Item("payrtot1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("bnstot1").Value = oRecordSet.Fields.Item("bnstot1").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("fwork1").Value = oRecordSet.Fields.Item("fwork1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = oRecordSet.Fields.Item("ndtalw1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("etcntax1").Value = oRecordSet.Fields.Item("etcntax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = oRecordSet.Fields.Item("lnchalw1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = oRecordSet.Fields.Item("ftaxamt1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = oRecordSet.Fields.Item("savtaxddc1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("incmtax1").Value = oRecordSet.Fields.Item("incmtax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = oRecordSet.Fields.Item("fvsptax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("residtax1").Value = oRecordSet.Fields.Item("residtax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = oRecordSet.Fields.Item("medcinsr1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = oRecordSet.Fields.Item("asopinsr1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = oRecordSet.Fields.Item("annuboamt1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("injungpay1").Value = oRecordSet.Fields.Item("injungpay1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("overpay1").Value = oRecordSet.Fields.Item("overpay1").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("entno2").Value = oRecordSet.Fields.Item("entno2").Value;
                                oForm.DataSources.UserDataSources.Item("servcomp2").Value = oRecordSet.Fields.Item("servcomp2").Value;
                                oForm.DataSources.UserDataSources.Item("symd2").Value = Convert.ToDateTime(oRecordSet.Fields.Item("symd2").Value.ToString().Trim()).ToString("yyyyMMdd");
                                oForm.DataSources.UserDataSources.Item("eymd2").Value = Convert.ToDateTime(oRecordSet.Fields.Item("eymd2").Value.ToString().Trim()).ToString("yyyyMMdd");

                                oForm.DataSources.UserDataSources.Item("payrtot2").Value = oRecordSet.Fields.Item("payrtot2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("bnstot2").Value = oRecordSet.Fields.Item("bnstot2").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("fwork2").Value = oRecordSet.Fields.Item("fwork2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = oRecordSet.Fields.Item("ndtalw2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("etcntax2").Value = oRecordSet.Fields.Item("etcntax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = oRecordSet.Fields.Item("lnchalw2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = oRecordSet.Fields.Item("ftaxamt2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = oRecordSet.Fields.Item("savtaxddc2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("indmtax2").Value = oRecordSet.Fields.Item("indmtax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = oRecordSet.Fields.Item("fvsptax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("residtax2").Value = oRecordSet.Fields.Item("residtax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = oRecordSet.Fields.Item("medcinsr2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = oRecordSet.Fields.Item("asopinsr2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = oRecordSet.Fields.Item("annuboamt2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("injungpay2").Value = oRecordSet.Fields.Item("injungpay2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("overpay2").Value = oRecordSet.Fields.Item("overpay2").Value.ToString();

                                // 부서
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                                sQry = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " and Code = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY401);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            PH_PY401_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY401_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY401_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY401_FormItemEnabled();
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
    }
}
