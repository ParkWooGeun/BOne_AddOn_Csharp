using System;
using System.IO;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using MsOutlook = Microsoft.Office.Interop.Outlook;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 원천징수영수증출력
    /// </summary>
    internal class PH_PY920 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY920A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY920B; //등록라인
        string SDocEntry;
        private string oLastItemUID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY920.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY920_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY920");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";
                oForm.Freeze(true);

                PH_PY920_CreateItems();
                PH_PY920_FormItemEnabled();
                PH_PY920_SetDocEntry();
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
        private void PH_PY920_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PH_PY920A = oForm.DataSources.DBDataSources.Item("@PH_PY920A");
                oDS_PH_PY920B = oForm.DataSources.DBDataSources.Item("@PH_PY920B");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
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

                // 출력구분
                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div").Specific.ValidValues.Add("1", "소득자보관용");
                oForm.Items.Item("Div").Specific.ValidValues.Add("2", "발행자보관용");
                oForm.Items.Item("Div").Specific.ValidValues.Add("3", "발행자보고용");
                oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 재직구분
                oForm.DataSources.UserDataSources.Add("Div1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div1").Specific.ValidValues.Add("1", "전체");
                oForm.Items.Item("Div1").Specific.ValidValues.Add("2", "재직자");
                oForm.Items.Item("Div1").Specific.ValidValues.Add("3", "퇴직자");
                oForm.Items.Item("Div1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 재직구분
                oForm.DataSources.UserDataSources.Add("Div2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div2").Specific.ValidValues.Add("1", "전체출력");
                oForm.Items.Item("Div2").Specific.ValidValues.Add("2", "첫장(1)만출력");
                oForm.Items.Item("Div2").Specific.ValidValues.Add("3", "첫장(2)만출력");
                oForm.Items.Item("Div2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        ///  <summary>
        ///  화면의 아이템 Enable 설정
        ///  </summary>
        private void PH_PY920_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YYYY").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    PH_PY920_SetDocEntry();
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YYYY").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("YYYY").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }


        /// <summary>
        /// PH_PY920_SetDocEntry
        /// </summary>
        private void PH_PY920_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY920'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY920A_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("YYYY").Specific.Value.Trim()))
                {
                    errMessage = "정산년도는 필수입니다.";
                    throw new System.Exception();
                }
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                    return returnValue;
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY920_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// PH_PY920_MTX01
        /// </summary>
        private void PH_PY920_MTX01()
        {
            int i;
            string sQry;
            string errMessage = string.Empty;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Param02 = oForm.Items.Item("YYYY").Specific.Value.Trim();
                Param03 = oForm.Items.Item("TeamCode").Specific.Value.Trim();
                Param04 = oForm.Items.Item("RspCode").Specific.Value.Trim();
                Param05 = oForm.Items.Item("ClsCode").Specific.Value.Trim();
                Param06 = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
                Param07 = oForm.Items.Item("Div1").Specific.Value.Trim();

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                sQry = "EXEC PH_PY920_99_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04+ "','" + Param05 + "','" + Param06 + "','" + Param07 + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PH_PY920B.Clear(); //추가

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "결과값이 존재하지않습니다.";
                    oMat01.Clear();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY920B.Size)
                    {
                        oDS_PH_PY920B.InsertRecord(i);
                    }
                    oMat01.AddRow();

                    oDS_PH_PY920B.Offset = i;
                    oDS_PH_PY920B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY920B.SetValue("U_sabun", i, oRecordSet.Fields.Item("sabun").Value);
                    oDS_PH_PY920B.SetValue("U_kname", i, oRecordSet.Fields.Item("u_fullname").Value);
                    oDS_PH_PY920B.SetValue("U_email", i, oRecordSet.Fields.Item("email").Value);
                    oDS_PH_PY920B.SetValue("U_StrDate", i, oRecordSet.Fields.Item("StrDate").Value);
                    oDS_PH_PY920B.SetValue("U_EndDate", i, oRecordSet.Fields.Item("EndDate").Value);
                    
                    if (!string.IsNullOrEmpty(oRecordSet.Fields.Item("email").Value.ToString().Trim()))
                    {
                        oDS_PH_PY920B.SetValue("U_Check", i, Convert.ToString('Y'));
                    }
                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    ProgressBar01.Stop();
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY920_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY920_Print_Report01()
        {
            string WinTitle;
            string ReportName = string.Empty;

            string CLTCOD;
            string YYYY;
            string TeamCode;
            string RspCode;
            string ClsCode;
            string MSTCOD;
            string Div;
            string Gubun;
            string Div1;

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
                Div = oForm.Items.Item("Div").Specific.Value.Trim();
                Gubun = oForm.Items.Item("Div2").Specific.Value.Trim();  //출력구분
                Div1 = oForm.Items.Item("Div1").Specific.Value.Trim();   //재직구분

                if (Convert.ToInt32(YYYY) >= 2023)
                {
                    //2023년귀속
                    WinTitle = "[PH_PY920] 원천징수영수증출력 2023년";

                    if (Gubun == "1")
                    {
                        ReportName = "PH_PY920_23_01.rpt";
                    }
                    else if (Gubun == "2")
                    {
                        ReportName = "PH_PY920_23_02.rpt";
                    }
                    else if (Gubun == "3")
                    {
                        ReportName = "PH_PY920_23_03.rpt";
                    }

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY)); //년
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", Div));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));
                    dataPackParameter.Add(new PSH_DataPackClass("@Div", Div1));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB41"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB53"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB61"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB62"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB63"));

                    if (Gubun == "1")
                    {
                        formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                    }
                    else
                    {
                        formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                    }
                }
                else
                {
                    //2022년귀속
                    WinTitle = "[PH_PY920] 원천징수영수증출력 2022년";

                    if (Gubun == "1")
                    {
                        ReportName = "PH_PY920_22_01.rpt";
                    }
                    else if (Gubun == "2")
                    {
                        ReportName = "PH_PY920_22_02.rpt";
                    }
                    else if (Gubun == "3")
                    {
                        ReportName = "PH_PY920_22_03.rpt";
                    }

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY)); //년
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", Div));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));
                    dataPackParameter.Add(new PSH_DataPackClass("@Div", Div1));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB41"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB41"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB53"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB53"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB61"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB61"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB62"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB62"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB63"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB63"));


                    if (Gubun == "1")
                    {
                        formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                    }
                    else
                    {
                        formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PDF만들기
        /// </summary>
        [STAThread]
        private bool Make_PDF_File(String p_MSTCOD)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName = String.Empty;
            string CLTCOD;
            string Main_Folder;
            string Sub_Folder1;
            string Sub_Folder2;
            string sQry;
            string ExportString;
            string psgovID;
            string YYYY;
            string TeamCode;
            string RspCode;
            string ClsCode;
            string MSTCOD;
            string Div;
            string Gubun;
            string Div1;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();
                MSTCOD = p_MSTCOD;
                Div = oForm.Items.Item("Div").Specific.Value.Trim();
                Gubun = oForm.Items.Item("Div2").Specific.Value.Trim();  //출력구분
                Div1 = oForm.Items.Item("Div1").Specific.Value.Trim();   //재직구분

                WinTitle = "[PH_PY920] 연말정산내역서";
                if (Gubun == "1")
                {
                    ReportName = "PH_PY920_22_01.rpt";
                }
                else if (Gubun == "2")
                {
                    ReportName = "PH_PY920_22_02.rpt";
                }
                else if (Gubun == "3")
                {
                    ReportName = "PH_PY920_22_03.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY)); //년
                dataPackFormula.Add(new PSH_DataPackClass("@Div", Div));

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));
                dataPackParameter.Add(new PSH_DataPackClass("@Div", Div1));

                //SubReport Parameter
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB1"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB1"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB1"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB1"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB1"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB1"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB11"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB11"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB11"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB11"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB11"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB11"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB2"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB2"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB2"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB2"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB2"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB2"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB21"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB21"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB21"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB21"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB21"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB21"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB3"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB3"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB3"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB3"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB3"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB3"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB4"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB4"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB4"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB4"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB4"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB4"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB41"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB41"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB41"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB41"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB41"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB41"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB5"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB5"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB5"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB5"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB5"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB5"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB51"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB51"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB51"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB51"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB51"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB51"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB52"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB52"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB52"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB52"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB52"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB52"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB53"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB53"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB53"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB53"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB53"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB53"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB61"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB61"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB61"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB61"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB61"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB61"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB62"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB62"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB62"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB62"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB62"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB62"));

                dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB63"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB63"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB63"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB63"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB63"));
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB63"));


                Main_Folder = @"C:\PSH_원천징수영수증";
                Sub_Folder1 = @"C:\PSH_원천징수영수증\" + YYYY + "";
                Sub_Folder2 = @"C:\PSH_원천징수영수증\" + YYYY + @"\" + CLTCOD + "";

                Dir_Exists(Main_Folder);
                Dir_Exists(Sub_Folder1);
                Dir_Exists(Sub_Folder2);

                ExportString = Sub_Folder2 + @"\" + p_MSTCOD + ".pdf";

                sQry = "Select RIGHT(U_govID,7) From [@PH_PY001A]";
                sQry += "WHERE  Code ='" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);
                psgovID = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                if (Gubun == "1")
                {
                    formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName, ExportString);
                }
                else
                {
                    formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula, ExportString, 100);
                }

                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(ExportString, PdfDocumentOpenMode.Modify);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                securitySettings.UserPassword = "manager";   //개개인암호
                securitySettings.OwnerPassword = psgovID;    //마스터암호

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // PDF문서 저장
                document.Save(ExportString);

                sQry = "Update [@PH_PY920B] Set U_SaveYN = 'Y' Where U_sabun = '" + p_MSTCOD + "' And DocEntry = '" + SDocEntry + "'";
                oRecordSet01.DoQuery(sQry);

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// 디렉토리 체크, 폴더 생성
        /// </summary>
        /// <param name="strDirName">경로</param>
        /// <returns></returns>
        private int Dir_Exists(string strDirName)
        {
            int ReturnValue = 0;

            try
            {
                DirectoryInfo di = new DirectoryInfo(strDirName); //DirectoryInfo 생성
                //DirectoryInfo.Exists로 폴더 존재유무 확인
                if (di.Exists)
                {
                    ReturnValue = 1;
                }
                else
                {
                    di.Create();
                    ReturnValue = 0;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Make_PDF_File_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }
        /// <summary>
        /// Send_EMail
        /// </summary>
        /// <param name="p_MSTCOD"></param>
        /// <param name="p_Version"></param>
        /// <returns></returns>
        private bool Send_EMail(string p_MSTCOD)
        {
            bool ReturnValue = false;
            string strToAddress;
            string strSubject;
            string strBody;
            string Sub_Folder2;
            string sQry;
            string YYYY;
            string MSTCOD;
            string CLTCOD;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MSTCOD = p_MSTCOD;
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();

                Sub_Folder2 = @"C:\PSH_원천징수영수증\" + YYYY + @"\" + CLTCOD + "";

                sQry = "Select U_Subject, U_Body From [@PH_PY920A] Where DocEntry = '" + SDocEntry + "'";
                oRecordSet01.DoQuery(sQry);
                strSubject = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                strBody = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                sQry = "SELECT U_eMail FROM [@PH_PY920B] WHERE U_sabun = '" + MSTCOD + "' AND DocEntry = '" + SDocEntry + "'";
                oRecordSet01.DoQuery(sQry);
                strToAddress = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                //mail.From = new MailAddress("dakkorea1@gmail.com");
                MsOutlook.Application outlookApp = new MsOutlook.Application();
                if (outlookApp == null)
                {
                    throw new Exception();
                }
                MsOutlook.MailItem mail = (MsOutlook.MailItem)outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);

                mail.Subject = strSubject;
                mail.HTMLBody = strBody;
                mail.To = strToAddress;
                MsOutlook.Attachment oAttach = mail.Attachments.Add(Sub_Folder2 + @"\" + p_MSTCOD + ".pdf");
                mail.Send();

                mail = null;
                outlookApp = null;

                sQry = "Update [@PH_PY920B] Set U_SendYN = 'Y' Where U_sabun = '" + p_MSTCOD + "' And DocEntry = '" + SDocEntry + "'";
                oRecordSet01.DoQuery(sQry);

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf");

                //원래코드시작
                //SmtpClient smtp = new SmtpClient("smtp.naver.com");
                //SmtpClient smtp = new SmtpClient("pscsn.poongsan.co.kr");
                //SmtpClient smtp = new SmtpClient("smtp.office365.com");
                //SmtpClient smtp = new SmtpClient("smtp.gmail.com");

                //smtp.Port = 587; //네이버
                //smtp.Port = 25; //풍산
                //smtp.UseDefaultCredentials = true;
                //smtp.EnableSsl = true;
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //smtp.Timeout = 20000;

                //smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;  //Naver 인 경우
                //smtp.Credentials = new NetworkCredential("2220501", "p2220501!"); //address, PW
                //smtp.Credentials = new NetworkCredential("wgpark@poongsan.co.kr", "1q2w3e4r)*"); //address, PW
                //smtp.Credentials = new NetworkCredential("dakkorea1@gmail.com", "dak440310*"); //address, PW

                //smtp.Send(mail);
                //원래코드 끝

                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Send_EMail_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return ReturnValue;
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
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                     //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                         //2
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                        //3
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                       //4
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                     //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK:                            //6
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                     //7
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:              //8
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:          //9
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:                         //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                      //11
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                  //12
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                        //16
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                      //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    //18
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                  //19
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                       //20
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                      //21
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    //22
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                //23
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                 //27
                    break;
                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                   //37
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                        //38
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag:                             //39
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
            string p_MSTCOD;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "btn_print")
                    {
                        if (PH_PY920A_DataValidCheck() == false)
                        {
                                BubbleEvent = false;
                                return;
                        }
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY920_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                    }
                    if (pVal.ItemUID == "btn_search")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY920A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oMat01.FlushToDataSource();
                            PH_PY920_MTX01();
                        }

                    }
                    //추가
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY920A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }

                            if (oDS_PH_PY920B.Size < 1)
                            {
                                errMessage = "조회 누르르고 추가하세오!";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "btn_save")
                    {
                        SDocEntry = oForm.Items.Item("DocEntry").Specific.Value;
                        if (PH_PY920A_DataValidCheck() == false)
                        {
                        }
                        else
                        {
                            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("PDF 파일 생성 시작!", 50, false);
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (!string.IsNullOrEmpty(oDS_PH_PY920B.GetValue("U_sabun", i).ToString().Trim()))
                                {
                                    if (oDS_PH_PY920B.GetValue("U_Check", i).ToString().Trim() == "Y")
                                    {
                                        p_MSTCOD = oDS_PH_PY920B.GetValue("U_sabun", i).ToString().Trim();
                                        if (Make_PDF_File(p_MSTCOD) == false)
                                        {
                                            errMessage = "PDF저장이 완료되지 않았습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                                ProgressBar01.Value += 1;
                                ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat01.VisualRowCount) + "건 PDF 파일 생성 중...!";
                            }
                            ProgressBar01.Stop();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PH_PY920_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = SDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                    if (pVal.ItemUID == "btn_send")
                    {
                        SDocEntry = oForm.Items.Item("DocEntry").Specific.Value;
                        if (PH_PY920A_DataValidCheck() == false)
                        {
                        }
                        else
                        {
                            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("PDF 파일 생성 시작!", 50, false);
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (oDS_PH_PY920B.GetValue("U_saveYN", i).ToString().Trim() == "Y")
                                {
                                    if (oDS_PH_PY920B.GetValue("U_Check", i).ToString().Trim() == "Y")
                                    {
                                        p_MSTCOD = oDS_PH_PY920B.GetValue("U_sabun", i).ToString().Trim();
                                        if (Send_EMail(p_MSTCOD) == false)
                                        {
                                            errMessage = "PDF저장이 완료되지 않았습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                                ProgressBar01.Value += 1;
                                ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat01.VisualRowCount) + "건 E-mail전송 중...!";
                            }
                            ProgressBar01.Stop();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PH_PY920_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = SDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY920_FormItemEnabled();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                        }
                    }
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
                            //사업장이 바뀌면 부서와 담당 재설정
                            case "CLTCOD":
                                //부서
                                //삭제
                                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                                //담당
                                //삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                //반
                                //삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            //부서가 바뀌면 담당 재설정
                            case "TeamCode":
                                //담당은 그 부서의 담당만 표시
                                //삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                //반
                                //삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            //담당이 바뀌면 반 재설정
                            case "RspCode":
                                //반
                                //삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                //현재 사업장으로 다시 Qry
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
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
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }
        /// <summary>
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = pVal.ColUID;
                        oLastColRow = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID = pVal.ItemUID;
                    oLastColUID = "";
                    oLastColRow = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY920_FormItemEnabled();
                }
            }
            catch (System.Exception ex)
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY920A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY920B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
