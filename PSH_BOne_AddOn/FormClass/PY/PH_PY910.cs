﻿using System;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 소득공제신고서출력
    /// </summary>
    internal class PH_PY910 : PSH_BaseClass
    {
        private string oFormUniqueID01;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY910.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY910_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY910");

                string strXml;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);

                PH_PY910_CreateItems();
                PH_PY910_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PH_PY910_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.DataBind.SetBound(true, "", "YYYY");
                oForm.DataSources.UserDataSources.Item("YYYY").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("MSTNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTNAME").Specific.DataBind.SetBound(true, "", "MSTNAME");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        ///  <summary>
        ///  화면의 아이템 Enable 설정
        ///  </summary>
        private void PH_PY910_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY910_Print_Report01()
        {
            string WinTitle;
            string ReportName;

            string CLTCOD;
            string YYYY;
            string TeamCode;
            string RspCode;
            string ClsCode;
            string MSTCOD;

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                YYYY = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (Convert.ToInt32(YYYY) >= 2022)
                {
                    //2023년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2023년";
                    ReportName = "PH_PY910_23_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB41"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB53"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB61"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB62"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB63"));

                    formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else
                {
                    //2022년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2022년";
                    ReportName = "PH_PY910_22_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB41"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB53"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB61"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB62"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB63"));

                    formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:            //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                //2
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:               //3
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:              //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:            //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK:                   //6
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:            //7
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:     //8
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE:                //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:             //11
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:         //12
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:               //16
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:             //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:           //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:         //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:              //20
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:             //21
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:           //22
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:       //23
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:        //27
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:          //37
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT:               //38
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag:                    //39
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY910_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        switch (pVal.ItemUID)
                        {
                            // 사업장이 바뀌면 부서와 담당 재설정
                            case "CLTCOD":
                                // 부서
                                // 삭제
                                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                // 현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                                // 담당
                                // 삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                // 반
                                // 삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                // 현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            // 부서가 바뀌면 담당 재설정
                            case "TeamCode":
                                // 담당은 그 부서의 담당만 표시
                                // 삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                // 현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                // 반
                                // 삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                // 현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            // 담당이 바뀌면 반 재설정
                            case "RspCode":
                                // 반
                                // 삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                // 현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim()() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("MSTNAME").Specific.Value = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
