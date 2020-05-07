//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;

using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여 E-Mail 발송 (2019.12.16 현재 사용안함, 사용하는 것으로 결정되었을 때 기능 테스트 필요, 송명규) 메일 발송 기능을 사용하게 되면, C#.NET 구문으로 신규 개발 필요
    /// </summary>
    internal class PH_PY118 : PSH_BaseClass
    {
        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;

        // 매트릭스 사용시
        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY118B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        private string oJOBYMM;
        private string oJOBTYP;
        private string oJOBGBN;
        private string oPAYSEL;
        private string oCLTCOD;
        private string oMSTBRK;
        private string oMSTDPT;
        private string oMSTCOD;

        private object sHtml;
        private string[] ArrPayHead = new string[37];
        private double[] ArrPayAmt = new double[37];
        private string[] ArrSubHead = new string[37];
        private double[] ArrSubAmt = new double[37];
        private string[] ArrGntHead = new string[19];
        private double[] ArrGntAmt = new double[19];
        private string sMSTCOD;
        private string sMSTNAM;
        private string sDPTNAM;
        private string sPOSITION;
        private double sTOTPAY;
        private double sTOTGON;
        private double sSILJIG;
        private string sTOEmail;
        private string sFrEmail;
        private string sFrSMTP;
        private string sFrSMTPSrv;
        private string sFrSMTPPort;
        private string sFrPWD;
        private string oPRTTIL;
        private string oPRTSUB;
        private string oCLTNAM;
        private string oREMARK;
        private bool oPrtChk;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {

            string strXml = string.Empty;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY118.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY118_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY118");

                strXml = oXmlDoc.xml.ToString();
                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                // oForm.DataBrowser.BrowseBy = "DocNum"

                oForm.Freeze(true);
                PH_PY118_CreateItems();
                PH_PY118_EnableMenus();
                PH_PY118_SetDocument(oFromDocEntry01);
            }
            catch(Exception ex)
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
        private void PH_PY118_CreateItems()
        {
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                //데이터셋(Matrix)
                oDS_PH_PY118B = oForm.DataSources.DBDataSources.Item("@PH_PY118B");

                //귀속년월
                oForm.DataSources.UserDataSources.Add("JOBYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("JOBYMM").Specific.DataBind.SetBound(true, "", "JOBYMM");
                oForm.DataSources.UserDataSources.Item("JOBYMM").Value = DateTime.Now.ToString("yyyyMM"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMM");

                //지급종류
                oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.Items.Item("JOBTYP").Specific.DataBind.SetBound(true, "", "JOBTYP");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                //지급구분
                oForm.DataSources.UserDataSources.Add("JOBGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.Items.Item("JOBGBN").Specific.DataBind.SetBound(true, "", "JOBGBN");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                //지급대상자구분
                oForm.DataSources.UserDataSources.Add("PAYSEL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.Items.Item("PAYSEL").Specific.DataBind.SetBound(true, "", "PAYSEL");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYSEL").Specific, "");
                oForm.Items.Item("PAYSEL").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PAYSEL").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("PAYSEL").DisplayDesc = true;

                //부서
                oForm.DataSources.UserDataSources.Add("MSTDPT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.Items.Item("MSTDPT").Specific.DataBind.SetBound(true, "", "MSTDPT");
                oForm.Items.Item("MSTDPT").DisplayDesc = true;

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //사원코드
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //사원명
                oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

                //보내는사람주소
                oForm.DataSources.UserDataSources.Add("FrEmail", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("FrEmail").Specific.DataBind.SetBound(true, "", "FrEmail");

                //SMTP server
                oForm.DataSources.UserDataSources.Add("FrSMTP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("FrSMTP").Specific.DataBind.SetBound(true, "", "FrSMTP");

                //패스워드
                oForm.DataSources.UserDataSources.Add("FrPWD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("FrPWD").Specific.DataBind.SetBound(true, "", "FrPWD");

                //공지사항
                oForm.DataSources.UserDataSources.Add("Remark", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                oForm.Items.Item("Remark").Specific.DataBind.SetBound(true, "", "Remark");

                //익명인증사용
                oForm.DataSources.UserDataSources.Add("AUTCHK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("AUTCHK").Specific.DataBind.SetBound(true, "", "AUTCHK");

                //6.기본설정 가져오기
                sQry = " SELECT U_FrSMTP, U_FrEMAIL, U_FrPWD FROM [@PH_PY118A] WHERE Code= '1'";
                oRecordSet.DoQuery(sQry);

                while (!(oRecordSet.EoF))
                {
                    oForm.DataSources.UserDataSources.Item("FrSMTP").Value = oRecordSet.Fields.Item(0).Value;
                    oForm.DataSources.UserDataSources.Item("FrEMAIL").Value = oRecordSet.Fields.Item(1).Value;
                    oForm.DataSources.UserDataSources.Item("FrPWD").Value = oRecordSet.Fields.Item(2).Value;
                    oRecordSet.MoveNext();
                }

                //Matrix
                oMat1 = oForm.Items.Item("Mat1").Specific;
                dataHelpClass.PAY_Matrix_AddCol(oMat1, "Col08", 121, "포함", 50, true, false, true, "@PH_PY118B", "U_Col08");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col01", 16, "부서", 80, true, false, true, "@PH_PY118B", "U_Col01");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col02", 16, "직책", 80, true, false, true, "@PH_PY118B", "U_Col02");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col03", 16, "입사일", 70, true, false, true, "@PH_PY118B", "U_Col03");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col04", 16, "퇴사일", 70, true, false, true, "@PH_PY118B", "U_Col04");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col05", 16, "실지급액", 80, true, true, true, "@PH_PY118B", "U_Col05");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col06", 16, "Email주소", 80, true, false, true, "@PH_PY118B", "U_Col06");
                dataHelpClass.PAY_Matrix_AddCol( oMat1, "Col07", 16, "확인", 50, true,  false, true, "@PH_PY118B", "U_Col07");

                //Check 버튼
                oForm.Items.Item("AUTCHK").Specific.ValOff = "N";
                oForm.Items.Item("AUTCHK").Specific.ValOn = "Y";

                oMat1.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 세팅(Enable)
        /// </summary>
        private void PH_PY118_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1282", false); //추가
                oForm.EnableMenu("1283", false); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 초기화(Set)
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY118_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY118_FormItemEnabled();
                    //PH_PY118_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY118_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 아이템 세팅(Enable)
        /// </summary>
        private void PH_PY118_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Items.Item("Btn2").Visible == true)
                {
                    oForm.Items.Item("FrSMTP").Visible = true;
                    oForm.Items.Item("FrEmail").Visible = true;
                    oForm.Items.Item("FrPWD").Visible = true;
                    oForm.Items.Item("s06").Visible = true;
                    oForm.Items.Item("s07").Visible = true;
                    oForm.Items.Item("s08").Visible = true;
                    oForm.Items.Item("Btn2").Visible = false;
                    oForm.Items.Item("Btn3").Visible = true;
                }
                else
                {
                    oForm.Items.Item("Remark").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("FrSMTP").Visible = false;
                    oForm.Items.Item("FrEmail").Visible = false;
                    oForm.Items.Item("FrPWD").Visible = false;
                    oForm.Items.Item("s06").Visible = false;
                    oForm.Items.Item("s07").Visible = false;
                    oForm.Items.Item("s08").Visible = false;
                    oForm.Items.Item("Btn2").Visible = true;
                    oForm.Items.Item("Btn3").Visible = false;
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
        /// SetSendMail
        /// </summary>
        private void SetSendMail()
        {
            string sQry = string.Empty;
            short errNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sFrEmail = oForm.Items.Item("FrEmail").Specific.Value.ToString().Trim();
                sFrSMTP = oForm.Items.Item("FrSMTP").Specific.Value.ToString().Trim();
                sFrPWD = oForm.Items.Item("FrPWD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(sFrEmail))
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(sFrSMTP))
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(sFrPWD))
                {
                    errNum = 3;
                    throw new Exception();
                }

                sQry = " SELECT U_FrSMTP, U_FrEMAIL, U_FrPWD FROM [@PH_PY118A] WHERE Code='1'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    sQry = "INSERT INTO [@PH_PY118A] (Code, Name, U_FrSMTP, U_FrEMAIL, U_FrPWD) values ('1','1','";
                    sQry = sQry + sFrSMTP + "', '" + sFrEmail + "', '" + sFrPWD + "')";
                    oRecordSet.DoQuery(sQry);
                }
                else
                {
                    sQry = "UPDATE  [@PH_PY118A] SET   U_FrSMTP = '" + sFrSMTP + "'";
                    sQry = sQry + " , U_FrEMAIL = '" + sFrEmail + "'";
                    sQry = sQry + " , U_FrPWD = '" + sFrPWD + "'";
                    sQry = sQry + " WHERE Code  = '1'";
                    oRecordSet.DoQuery(sQry);
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("보내는 사람주소는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("SMTP Server는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PassWord는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PrintChk
        /// </summary>
        private void PrintChk()
        {
            short i = 0;

            try
            {
                if (oMat1.RowCount == 1)
                {
                    return;
                }

                oMat1.FlushToDataSource();
                for (i = 0; i <= oDS_PH_PY118B.Size - 1; i++)
                {
                    oDS_PH_PY118B.Offset = i;
                    if (oPrtChk == true)
                    {
                        oDS_PH_PY118B.SetValue("U_Col11", i, "N");
                    }
                    else
                    {
                        oDS_PH_PY118B.SetValue("U_Col11", i, "Y");
                    }
                }
                oMat1.LoadFromDataSource();
                if (oPrtChk == true)
                {
                    oPrtChk = false;
                }
                else
                {
                    oPrtChk = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Create_Html2
        /// </summary>
        private void Create_Html2()
        {
            short i = 0;

            try
            {
                sHtml = "<html>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<head>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<title>급여명세서 e-MAIL발송</title>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<style type=\"text/css\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<!--" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + " td {  font-size: 9pt; line-height: 14pt; color: #000000}" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + ".목록 {  font-size: 9pt; font-weight: bold; color: #FFFFFF}" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + ".마침글 {  font-size: 9pt; color: #000000; font-weight: bold}}" + Environment.NewLine; //Constants.vbCrLf;
                //파란색:#3333FF
                sHtml = sHtml + "-->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</style>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</head>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<body>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</style>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</head>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "<body>" + Environment.NewLine; //Constants.vbCrLf;
                //BODY /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                //수정0->1
                sHtml = sHtml + " <table border=\"0\" cellspacing=\"0\" cellpadding=\"0\"  align=\"center\" width=\"900\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "      <tr></tr>" + Environment.NewLine; //Constants.vbCrLf;
                //**********************************************************************************************************************/
                //타이틀정보
                //**********************************************************************************************************************/
                sHtml = sHtml + "     <table border=\"0\" bgcolor=\"white\" cellspacing=\"0\" cellpadding=\"2\" width=\"800\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      <tr bgcolor=\"white\" align=\"Center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"100%\"><h4><b>&lt;" + oPRTTIL + "&gt;</h></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      <tr bgcolor=\"white\" align=\"Center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"100%\"><h4><b>" + oPRTSUB + "</h></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "     </table><br>" + Environment.NewLine; //Constants.vbCrLf;
                //**********************************************************************************************************************/
                //사원정보
                //**********************************************************************************************************************/
                sHtml = sHtml + "     <table border=\"1\" bgcolor=\"white\" cellspacing=\"0\" cellpadding=\"2\" width=\"800\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"12%\">사  번</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"13%\">성  명</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"15%\">부  서</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"15%\">직  책</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"15%\">지급총액</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"15%\">공제총액</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"15%\">실지급액</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"CENTER\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"12%\" align=\"CENTER\">" + sMSTCOD + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"13%\" align=\"CENTER\">" + sMSTNAM + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"15%\" align=\"CENTER\">" + sDPTNAM + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"15%\" align=\"CENTER\">" + sPOSITION + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"15%\" align=\"Right\"><b>" + sTOTPAY.ToString("#,##0") + "</b></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"15%\" align=\"Right\"><b>" + sTOTGON.ToString("#,##0") + "</b></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"15%\" align=\"Right\"><b>" + sSILJIG.ToString("#,##0") + "</b></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "     </table><br>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "  <tr>" + Environment.NewLine; //Constants.vbCrLf;
                //**********************************************************************************************************************/
                //내역
                //**********************************************************************************************************************/
                sHtml = sHtml + "                <table border=\"1\"  bgcolor=\"white\" cellspacing=\"0\" cellpadding=\"2\" width=\"800\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                <!--- 근태항목 ----------------------------------------------------------->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"10%\" rowspan=\"4\">근태항목</td>" + Environment.NewLine; //Constants.vbCrLf;
                //근태항목(01~09항목)
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td  width=\"10%\">" + (ArrGntHead[i].Trim() == "" ? "-" : ArrGntHead[i].Trim()) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrGntAmt[i] + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //근태항목(10~18항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td  width=\"10%\">" + (ArrGntHead[i].Trim() == "" ? "-" : ArrGntHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrGntAmt[i] + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;

                sHtml = sHtml + "                <!--- 지급항목 ----------------------------------------------------------->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td  width=\"10%\" rowspan=\"8\">지급항목</td>" + Environment.NewLine; //Constants.vbCrLf;
                //지급항목(01~09항목)
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td  width=\"10%\">" + ArrPayHead[i] + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrPayAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //지급항목(10~18항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrPayHead[i].Trim() == "" ? "-" : ArrPayHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrPayAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //지급항목(19~27항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 19; i <= 27; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrPayHead[i].Trim() == "" ? "-" : ArrPayHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 19; i <= 27; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrPayAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //지급항목(28~36항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 28; i <= 36; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrPayHead[i].Trim() == "" ? "-" : ArrPayHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 28; i <= 36; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrPayAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;

                sHtml = sHtml + "<!--- 공제항목 ----------------------------------------------------------->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                        <td width=\"10%\" rowspan=\"8\">공제항목</td>" + Environment.NewLine; //Constants.vbCrLf;
                //공제항목(01-09항목)
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td  width=\"10%\">" + (ArrSubHead[i].Trim() == "" ? "-" : ArrSubHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 1; i <= 9; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrSubAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //공제항목(08-14항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                //14
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrSubHead[i].Trim() == "" ? "-" : ArrSubHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                //14
                for (i = 10; i <= 18; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrSubAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //공제항목(08-14항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 19; i <= 27; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrSubHead[i].Trim() == "" ? "-" : ArrSubHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                for (i = 19; i <= 27; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrSubAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                //공제항목(08-14항목)
                sHtml = sHtml + "                    <tr bgcolor=\"RGB(239,235,222)\" align=\"center\">" + Environment.NewLine; //Constants.vbCrLf;
                //14
                for (i = 28; i <= 36; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + (ArrSubHead[i].Trim() == "" ? "-" : ArrSubHead[i]) + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                    <tr bgcolor=\"white\" align=\"Right\">" + Environment.NewLine; //Constants.vbCrLf;
                //14
                for (i = 28; i <= 36; i++)
                {
                    sHtml = sHtml + "                        <td width=\"10%\">" + ArrSubAmt[i].ToString("#,##0") + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                }
                sHtml = sHtml + "                    </tr>" + Environment.NewLine; //Constants.vbCrLf;

                sHtml = sHtml + "                </table>" + Environment.NewLine; //Constants.vbCrLf;
                //**********************************************************************************************************************/
                //비고내용
                //**********************************************************************************************************************/
                sHtml = sHtml + "     <table border=\"0\" bgcolor=\"white\" cellspacing=\"0\" cellpadding=\"2\" width=\"800\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      <tr bgcolor=\"white\" align=\"Left\">" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                          <td width=\"100%\">" + oREMARK + "</td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "                      </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "     </table><br>" + Environment.NewLine; //Constants.vbCrLf;

                sHtml = sHtml + "<!--- 지급항목테이블 끝. ----------------------------------------------------------------------->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "      </td></tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "      <tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "       <td align=\"left\" ><b class=\"마침글\">" + oCLTNAM + "</b></td>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "      </tr>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + " </table>" + Environment.NewLine; //Constants.vbCrLf;

                //END/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                sHtml = sHtml + "<!--- 본문 끝입니다. -------------------------------------------------------------------------->" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</body>" + Environment.NewLine; //Constants.vbCrLf;
                sHtml = sHtml + "</html>" + Environment.NewLine; //Constants.vbCrLf;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Send_eMail
        /// 메일 발송 기능을 사용하게 되면, C#.NET 구문으로 신규 개발 필요
        /// </summary>
        /// <returns></returns>
        private string Send_eMail()
        {
            string returnValue = string.Empty;

            //object MailSender = Activator.CreateInstance(Type.GetTypeFromProgID("CDO.Message")); //CDO.Message 개체에 대한 참조를 작성하고 반환
            //object iConf = Activator.CreateInstance(Type.GetTypeFromProgID("CDO.Configuration")); //CDO.Configuration개체에 대한 참조를 작성하고 반환

            ////MailSender = Activator.CreateInstance(Type.GetTypeFromProgID("CDO.Message")); //Interaction.CreateObject("CDO.Message");
            ////iConf = Activator.CreateInstance(Type.GetTypeFromProgID("CDO.Configuration")); //Interaction.CreateObject("CDO.Configuration");

            //try
            //{
            //    if (oForm.Items.Item("AUTCHK").Specific.Checked == false)
            //    {

            //        var _with1 = iConf.Fields;

            //        _with1.Refresh();
            //        _with1.Item(CDO.CdoConfiguration.cdoSendUsingMethod).Value = CDO.CdoSendUsing.cdoSendUsingPort;
            //        //1:로컬smtp로 메일전송, 2 cdoSendUsingPort: 외부smtp로 메일전송
            //        _with1.Item(CDO.CdoConfiguration.cdoSMTPServer).Value = sFrSMTPSrv;
            //        //"mail.emdc.co.kr","mail.care-line.co.kr" '"보내는 사람의 SMTP Server Name" '(ex: mail.xxx.com)
            //        _with1.Item(CDO.CdoConfiguration.cdoSMTPConnectionTimeout).Value = 10;
            //        _with1.Item(CDO.CdoConfiguration.cdoSMTPAuthenticate).Value = CDO.CdoProtocolsAuthentication.cdoBasic;
            //        //기본인증
            //        _with1.Item(CDO.CdoConfiguration.cdoSendUserName).Value = sFrEmail;
            //        //"hammi97@emdc.co.kr""sap2@care-line.co.kr" '"보내는사람주소" '(ex : peter@xxx.com)
            //        _with1.Item(CDO.CdoConfiguration.cdoSendPassword).Value = sFrPWD;
            //        //mi0215" ''"sap02" '"보내는 사람의 Password"
            //        _with1.Item(CDO.CdoConfiguration.cdoURLGetLatestVersion).Value = true;
            //        _with1.Item(CDO.CdoConfiguration.cdoSMTPServerPort).Value = sFrSMTPPort;
            //        //.Item(cdoSMTPServerPort) = 25 '/통상 25번포트 네이버 pop3:110포트

            //        _with1.Update();
            //    }
            //    else
            //    {
            //        var _with2 = iConf.Fields;

            //        _with2.Refresh();
            //        _with2.Item(CDO.CdoConfiguration.cdoSendUsingMethod).Value = CDO.CdoSendUsing.cdoSendUsingPort;
            //        ///1:로컬smtp로 메일전송, 2 cdoSendUsingPort: 외부smtp로 메일전송
            //        _with2.Item(CDO.CdoConfiguration.cdoSMTPServer).Value = sFrSMTPSrv;
            //        //"mail.care-line.co.kr" '"보내는 사람의 SMTP Server Name" '(ex: mail.xxx.com)
            //        _with2.Item(CDO.CdoConfiguration.cdoSMTPConnectionTimeout).Value = 10;
            //        _with2.Item(CDO.CdoConfiguration.cdoSMTPAuthenticate).Value = CDO.CdoProtocolsAuthentication.cdoAnonymous;
            //        //익명인증
            //        _with2.Item(CDO.CdoConfiguration.cdoURLGetLatestVersion).Value = true;
            //        _with2.Item(CDO.CdoConfiguration.cdoSMTPServerPort).Value = sFrSMTPPort;
            //        //.Item(cdoSMTPServerPort) = 25 '/통상 25번포트 네이버 pop3:110포트

            //        _with2.Update();
            //    }

            //    var _with3 = MailSender;
            //    _with3.let_Configuration(iConf);
            //    _with3.From = sFrEmail;
            //    //전송자/수신자 이름만 넘김
            //    _with3.To = sTOEmail;
            //    _with3.Subject = oPRTTIL;
            //    //UPGRADE_WARNING: sHtml 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            //    _with3.HTMLBody = sHtml;
            //    // .TextBody = "귀하의 노고에 감사드립니다."
            //    _with3.BodyPart.Charset = "ks_c_5601-1987";
            //    _with3.HTMLBodyPart.Charset = "ks_c_5601-1987";
            //    _with3.send();
            //    //.AddAttachment "C:\files\mybook.doc"   '/ 첨부파일

            //    MailSender = null;
            //    iConf = null;

            //    PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            //    returnValue = "True";
            //}
            //catch (Exception ex)
            //{
            //    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            //}
            //finally
            //{
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(MailSender);
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(iConf);
            //}

            return returnValue;
        }

        /// <summary>
        /// Execution_Process
        /// </summary>
        /// <returns></returns>
        private bool Execution_Process()
        {
            bool functionReturnValue = false;
            
            string sQry = string.Empty;
            short errNum = 0;
            int i = 0;
            int TOTCNT = 0;
            int V_StatusCnt = 0;
            int oProValue = 0;
            int tRow = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar oProgBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);

            try
            {
                //Check
                oJOBYMM = oForm.Items.Item("JOBYMM").Specific.Value.ToString().Trim();
                oJOBTYP = oForm.Items.Item("JOBTYP").Specific.Selected.Value.ToString().Trim();
                oJOBGBN = oForm.Items.Item("JOBGBN").Specific.Selected.Value.ToString().Trim();
                oPAYSEL = oForm.Items.Item("PAYSEL").Specific.Selected.Value.ToString().Trim();
                oCLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                oMSTDPT = oForm.Items.Item("MSTDPT").Specific.Selected.Value.ToString().Trim();
                oMSTCOD = oForm.Items.Item("MSTCOD").Specific.String;

                if (string.IsNullOrEmpty(oMSTCOD))
                {
                    oMSTCOD = "%";
                }

                if (string.IsNullOrEmpty(oJOBYMM))
                {
                    errNum = 1;
                    throw new Exception();
                }

                //switch (true)
                //{
                //    case string.IsNullOrEmpty(Strings.Trim(oJOBYMM)):
                //        errNum = 1;
                //        goto Error_Message;
                //        break;
                //}

                oDS_PH_PY118B.Clear();
                oMat1.LoadFromDataSource();
                //i = 0;
                sQry = "SELECT T0.U_MSTCOD,T0.U_MSTNAM, T0.U_EmpID,  T2.U_CodeNM, T3.Name,";
                sQry = sQry + " ISNULL(CONVERT(CHAR(10), T1.U_StartDat, 20),'') AS U_INPDAT,";
                sQry = sQry + " ISNULL(CONVERT(CHAR(10),T1.U_TermDate, 20), '') AS U_OUTDAT, T0.U_SILJIG, T1.U_email";
                sQry = sQry + " FROM [@PH_PY112A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code";
                sQry = sQry + " INNER JOIN [@PS_HR200L] T2 ON T1.U_TeamCode = T2.U_Code AND T2.Code = '1'";
                sQry = sQry + " INNER JOIN [OHPS] T3 ON T1.U_Position = T3.posID";
                sQry = sQry + " WHERE   T0.U_YM = '" + oJOBYMM + "'";
                sQry = sQry + " AND     T0.U_JOBTYP = '" + oJOBTYP + "'";
                sQry = sQry + " AND     T0.U_JOBGBN = '" + oJOBGBN + "'";
                sQry = sQry + " AND     (T1.U_PAYSEL = '" + oPAYSEL + "' OR T1.U_PAYSEL LIKE '" + oPAYSEL + "')";
                sQry = sQry + " AND     T0.U_CLTCOD = '" + oCLTCOD + "'";
                sQry = sQry + " AND     (T1.U_TeamCode = '" + oMSTDPT + "' OR T1.U_TeamCode LIKE '" + oMSTDPT + "')";
                sQry = sQry + " AND     (T1.Code = '" + oMSTCOD + "' OR T1.Code LIKE '" + oMSTCOD + "')";
                sQry = sQry + " ORDER BY T0.U_CLTCOD,  T0.U_TeamCode, T1.U_Position, T0.U_MSTCOD";

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                }

                if (oProgBar != null)
                {
                    oProgBar.Stop();
                }
                
                //최대값 구하기
                TOTCNT = oRecordSet.RecordCount;

                V_StatusCnt = TOTCNT / 50;
                oProValue = 1;
                tRow = 1;
                
                while (!(oRecordSet.EoF))
                {
                    oDS_PH_PY118B.InsertRecord((i));
                    oDS_PH_PY118B.Offset = i;
                    oDS_PH_PY118B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY118B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY118B.SetValue("U_MSTNAM", i, oRecordSet.Fields.Item(1).Value);
                    oDS_PH_PY118B.SetValue("U_EMPID", i, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_PY118B.SetValue("U_Col01", i, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_PY118B.SetValue("U_Col02", i, oRecordSet.Fields.Item(4).Value);
                    oDS_PH_PY118B.SetValue("U_Col03", i, oRecordSet.Fields.Item(5).Value);
                    oDS_PH_PY118B.SetValue("U_Col04", i, oRecordSet.Fields.Item(6).Value);
                    //oDS_PH_PY118B.SetValue("U_Col05", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item(7).Value, "#,###,###,##0"));
                    oDS_PH_PY118B.SetValue("U_Col05", i, oRecordSet.Fields.Item(7).Value.ToString());
                    oDS_PH_PY118B.SetValue("U_Col06", i, oRecordSet.Fields.Item(8).Value);
                    if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(8).Value.ToString().Trim()))
                    {
                        oDS_PH_PY118B.SetValue("U_Col08", i, "Y");
                    }
                    else
                    {
                        oDS_PH_PY118B.SetValue("U_Col08", i, "N");
                    }
                    i = i + 1;
                    oRecordSet.MoveNext();

                    if ((TOTCNT > 50 && tRow == oProValue * V_StatusCnt) || TOTCNT <= 50)
                    {
                        oProgBar.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                        oProValue = oProValue + 1;
                        oProgBar.Value = oProValue;
                    }
                    tRow = tRow + 1;
                }
                oPrtChk = true;
                oProgBar.Stop();
                oMat1.LoadFromDataSource();

                //End
                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                functionReturnValue = false;
                
                if (oProgBar != null)
                {
                    oProgBar.Stop();
                }
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속연월을 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조건과 일치하는 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    //PSH_Globals.SBO_Application.StatusBar.SetText("Execution_Process 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// eMail_Process
        /// </summary>
        private void eMail_Process()
        {
            string sQry = string.Empty;
            short errNum = 0;
            int i = 0;
            int cnt = 0;
            int oRow = 0;
            string RetVal = string.Empty;
            string[] GNTSTR = new string[10];

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                //Check
                oMat1.FlushToDataSource();
                sFrEmail = oForm.Items.Item("FrEmail").Specific.Value.ToString().Trim();
                sFrSMTP = oForm.Items.Item("FrSMTP").Specific.Value.ToString().Trim();
                sFrPWD = oForm.Items.Item("FrPWD").Specific.Value.ToString().Trim();
                oREMARK = oForm.Items.Item("Remark").Specific.Value.ToString().Trim();

                //VB(InStr)와 C#(IndexOf) 문법 차이
                //InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare)
                //BLOCKTAGLIST.IndexOf(";" + strTagName + ";", System.StringComparison.OrdinalIgnoreCase) + 1;

                if ((sFrSMTP.IndexOf(":", System.StringComparison.OrdinalIgnoreCase) + 1) == 0)
                {
                    sFrSMTPSrv = sFrSMTP;
                    sFrSMTPPort = "25";
                }
                else
                {
                    sFrSMTPSrv = codeHelpClass.Left(sFrSMTP, sFrSMTP.IndexOf(":", System.StringComparison.OrdinalIgnoreCase));
                    sFrSMTPPort = codeHelpClass.Mid(sFrSMTP, sFrSMTP.IndexOf(":", System.StringComparison.OrdinalIgnoreCase), sFrSMTP.Length - (sFrSMTP.IndexOf(":", System.StringComparison.OrdinalIgnoreCase) + 1));
                }

                if (oMat1.RowCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(sFrEmail))
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(sFrSMTP))
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(sFrPWD))
                {
                    errNum = 5;
                    throw new Exception();
                }

                //switch (true)
                //{
                //    case oMat1.RowCount == 0:
                //        errNum = 1;
                //        goto Error_Message;
                //        break;
                //    case string.IsNullOrEmpty(Strings.Trim(sFrEmail)):
                //        errNum = 3;
                //        goto Error_Message;
                //        break;
                //    case string.IsNullOrEmpty(Strings.Trim(sFrSMTP)):
                //        errNum = 4;
                //        goto Error_Message;
                //        break;
                //    case string.IsNullOrEmpty(Strings.Trim(sFrPWD)):
                //        errNum = 5;
                //        goto Error_Message;
                //        break;
                //}

                //초기화
                for (i = 1; i <= 36; i++)
                {
                    ArrPayHead[i] = "---";
                    ArrPayAmt[i] = 0;
                }
                for (i = 1; i <= 36; i++)
                {
                    ArrSubHead[i] = "---";
                    ArrSubAmt[i] = 0;
                }
                for (i = 1; i <= 18; i++)
                {
                    ArrGntHead[i] = "---";
                    ArrGntAmt[i] = 0;
                }

                //1. 수당/공제/근태 항목
                //수당 항목
                sQry = "SELECT T0.U_CSUNAM";
                sQry = sQry + " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE U_CLTCOD = '" + oCLTCOD + "'";
                sQry = sQry + " AND (T1.U_YM = '" + oJOBYMM + "' OR (T1.U_YM <> '" + oJOBYMM + "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY102A] WHERE U_YM <= '" + oJOBYMM + "' )))";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                }
                else
                {
                    for (i = 1; i <= 36; i++)
                    {
                        if (i <= oRecordSet.RecordCount)
                        {
                            ArrPayHead[i] = oRecordSet.Fields.Item(0).Value;
                            oRecordSet.MoveNext();
                        }
                    }
                }

                //공제 항목
                sQry = "SELECT T0.U_CSUNAM";
                sQry = sQry + " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE U_CLTCOD = '" + oCLTCOD + "'";
                sQry = sQry + " AND (T1.U_YM = '" + oJOBYMM + "' OR (T1.U_YM <> '" + oJOBYMM + "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY103A] WHERE U_YM <= '" + oJOBYMM + "' )))";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                }
                else
                {
                    for (i = 1; i <= 36; i++)
                    {
                        if (i <= oRecordSet.RecordCount)
                        {
                            ArrSubHead[i] = oRecordSet.Fields.Item(0).Value;
                            oRecordSet.MoveNext();
                        }
                    }
                }

                //2. 근태관련(1~14)
                ArrGntHead[1] = "근로일수";
                ArrGntHead[2] = "특근일수";
                ArrGntHead[3] = "유급일수";
                ArrGntHead[4] = "연장근로시간";
                ArrGntHead[5] = "야간근로";
                ArrGntHead[6] = "휴일기본";
                ArrGntHead[7] = "년월차";
                ArrGntHead[8] = "교대(주)";
                ArrGntHead[9] = "교대(야)";
                ArrGntHead[10] = "위해일수";
                ArrGntHead[11] = "생휴발생";
                ArrGntHead[12] = "생휴사용";
                ArrGntHead[13] = "지급율(%)";
                ArrGntHead[14] = "부양가족";
                ArrGntHead[15] = "일당액";
                ArrGntHead[16] = "---";
                ArrGntHead[17] = "---";
                ArrGntHead[18] = "---";

                //3. 전체 적용사항
                oCLTNAM = PSH_Globals.oCompany.CompanyName;
                switch (oJOBTYP)
                {
                    case "1":
                        oPRTTIL = codeHelpClass.Left(oJOBYMM, 4) + "년 " + codeHelpClass.Mid(oJOBYMM, 4, 2) + "월 급여 명세서";
                        break;
                    case "2":
                        oPRTTIL = codeHelpClass.Left(oJOBYMM, 4) + "년 " + codeHelpClass.Mid(oJOBYMM, 4, 2) + "월 상여 명세서";
                        break;
                }

                oPRTSUB = "(  " + oForm.Items.Item("JOBGBN").Specific.Selected.Description + "  )";
                //4.사원별 DM발송
                cnt = 0;
                for (oRow = 0; oRow <= oDS_PH_PY118B.Size - 1; oRow++)
                {
                    oDS_PH_PY118B.Offset = oRow;
                    if (oDS_PH_PY118B.GetValue("U_Col08", oRow).Trim() == "Y")
                    {
                        sMSTCOD = oDS_PH_PY118B.GetValue("U_MSTCOD", oRow).Trim();
                        sMSTNAM = oDS_PH_PY118B.GetValue("U_MSTNAM", oRow).Trim();
                        sDPTNAM = oDS_PH_PY118B.GetValue("U_Col01", oRow).Trim();
                        sPOSITION = oDS_PH_PY118B.GetValue("U_Col02", oRow).Trim();
                        sTOEmail = oDS_PH_PY118B.GetValue("U_Col06", oRow).Trim();

                        if (string.IsNullOrEmpty(sTOEmail))
                        {
                            oDS_PH_PY118B.SetValue("U_Col07", oRow, "메일주소누락");
                        }
                        else
                        {
                            //급여정보 가져오기
                            sQry = "SELECT  T0.U_MSTCOD,";
                            sQry = sQry + " T0.U_CSUD01, T0.U_CSUD02, T0.U_CSUD03, T0.U_CSUD04, T0.U_CSUD05, T0.U_CSUD06, T0.U_CSUD07, T0.U_CSUD08,T0.U_CSUD09,";
                            sQry = sQry + " T0.U_CSUD10, T0.U_CSUD11, T0.U_CSUD12, T0.U_CSUD13, T0.U_CSUD14, T0.U_CSUD15, T0.U_CSUD16,T0.U_CSUD17,T0.U_CSUD18,";
                            sQry = sQry + " T0.U_CSUD19, T0.U_CSUD20, T0.U_CSUD21, T0.U_CSUD22, T0.U_CSUD23, T0.U_CSUD24, T0.U_CSUD25, T0.U_CSUD26, T0.U_CSUD27,";
                            sQry = sQry + " T0.U_CSUD28, T0.U_CSUD29, T0.U_CSUD30, T0.U_CSUD31, T0.U_CSUD32, T0.U_CSUD33, T0.U_CSUD34, T0.U_CSUD35, T0.U_CSUD36,";
                            sQry = sQry + " T0.U_GONG01, T0.U_GONG02, T0.U_GONG03, T0.U_GONG04, T0.U_GONG05, T0.U_GONG06, T0.U_GONG07, T0.U_GONG08, T0.U_GONG09,";
                            sQry = sQry + " T0.U_GONG10, T0.U_GONG11, T0.U_GONG12, T0.U_GONG13, T0.U_GONG14, T0.U_GONG15, T0.U_GONG16, T0.U_GONG17, T0.U_GONG18,";
                            sQry = sQry + " T0.U_GONG19, T0.U_GONG20, T0.U_GONG21, T0.U_GONG22, T0.U_GONG23, T0.U_GONG24, T0.U_GONG25, T0.U_GONG26, T0.U_GONG27,";
                            sQry = sQry + " T0.U_GONG28, T0.U_GONG29, T0.U_GONG30, T0.U_GONG31, T0.U_GONG32, T0.U_GONG33, T0.U_GONG34, T0.U_GONG35, T0.U_GONG36,";
                            sQry = sQry + " T0.U_TOTPAY, T0.U_TOTGON, T0.U_SILJIG, T0.U_CLTNAM,";
                            sQry = sQry + " T3.U_GetDay, T3.U_WoHDay, T3.U_PayDay, T3.U_Extend, T3.U_Midnight, T3.U_Special, T3.U_YCHHGA,";
                            sQry = sQry + " T3.U_EtcDAY1 , T3.U_EtcDAY2, T3.U_WHMDAY, T3.U_SNHDAY, T3.U_SNHHGA, T0.U_APPRAT, T0.U_BUYNSU";
                            sQry = sQry + " FROM [@PH_PY112A] T0";
                            sQry = sQry + " LEFT JOIN ( SELECT  T2.U_MSTCOD, T1.U_YM, T2.U_GetDay, T2.U_WoHDay, T2.U_PayDay, T2.U_Extend, T2.U_Midnight, T2.U_Special,";
                            sQry = sQry + "                     T2.U_YCHHGA, T2.U_EtcDAY1 , T2.U_EtcDAY2, T2.U_WHMDAY, T2.U_SNHDAY, T2.U_SNHHGA";
                            sQry = sQry + "             FROM [@PH_PY017B] T2 INNER JOIN [@PH_PY017A] T1 ON T2.Code = T1.Code";
                            sQry = sQry + "           ) T3 ON T0.U_YM = T3.U_YM AND T0.U_MSTCOD = T3.U_MSTCOD";
                            sQry = sQry + " WHERE   T0.U_YM = '" + oJOBYMM + "'";
                            sQry = sQry + " AND     T0.U_JOBTYP = '" + oJOBTYP + "'";
                            sQry = sQry + " AND     T0.U_JOBGBN = '" + oJOBGBN + "'";
                            sQry = sQry + " AND     (T0.U_JOBTRG = '" + oPAYSEL + "' OR ( T0.U_JOBTRG <> '" + oPAYSEL + "' AND T0.U_JOBTRG LIKE '" + oPAYSEL + "'))";
                            sQry = sQry + " AND     T0.U_MSTCOD = '" + sMSTCOD + "'";

                            oRecordSet.DoQuery(sQry);
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (i = 1; i <= 36; i++)
                                {
                                    //if (object.ReferenceEquals(oRecordSet.Fields.Item("U_CSUD" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value, System.DBNull.Value))
                                    if (object.ReferenceEquals(oRecordSet.Fields.Item("U_CSUD" + i.ToString().PadLeft(2, '0')).Value, System.DBNull.Value))
                                    {
                                        //oRecordSet.Fields.Item("U_CSUD" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value = 0;
                                        oRecordSet.Fields.Item("U_CSUD" + i.ToString().PadLeft(2, '0')).Value = 0;
                                    }
                                    //ArrPayAmt[i] = oRecordSet.Fields.Item("U_CSUD" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value;
                                    ArrPayAmt[i] = oRecordSet.Fields.Item("U_CSUD" + i.ToString().PadLeft(2, '0')).Value;
                                }

                                for (i = 1; i <= 36; i++)
                                {
                                    //if (object.ReferenceEquals(oRecordSet.Fields.Item("U_GONG" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value, System.DBNull.Value))
                                    if (object.ReferenceEquals(oRecordSet.Fields.Item("U_GONG" + i.ToString().PadLeft(2, '0')).Value, System.DBNull.Value))
                                    {
                                        //oRecordSet.Fields.Item("U_GONG" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value = 0;
                                        oRecordSet.Fields.Item("U_GONG" + i.ToString().PadLeft(2, '0')).Value = 0;
                                    }
                                    //ArrSubAmt[i] = oRecordSet.Fields.Item("U_GONG" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(i, "00")).Value;
                                    ArrSubAmt[i] = oRecordSet.Fields.Item("U_GONG" + i.ToString().PadLeft(2, '0')).Value;
                                }

                                for (i = 1; i <= 14; i++)
                                {
                                    if (object.ReferenceEquals(oRecordSet.Fields.Item(i + 76).Value, System.DBNull.Value))
                                    {
                                        oRecordSet.Fields.Item(i + 76).Value = 0;
                                    }
                                    ArrGntAmt[i] = oRecordSet.Fields.Item(i + 76).Value;
                                }

                                sTOTPAY = oRecordSet.Fields.Item("U_TOTPAY").Value.ToString().Trim();
                                sTOTGON = oRecordSet.Fields.Item("U_TOTGON").Value.ToString().Trim();
                                sSILJIG = oRecordSet.Fields.Item("U_SILJIG").Value.ToString().Trim();

                                if (!string.IsNullOrEmpty(oRecordSet.Fields.Item("U_CLTNAM").Value))
                                {
                                    oCLTNAM = oRecordSet.Fields.Item("U_CLTNAM").Value;
                                }
                            }
                            else
                            {
                                oDS_PH_PY118B.SetValue("U_Col07", oRow, "자료누락.");

                                for (i = 1; i <= 36; i++)
                                {
                                    ArrPayAmt[i] = 0;
                                }
                                for (i = 1; i <= 36; i++)
                                {
                                    ArrSubAmt[i] = 0;
                                }
                                for (i = 1; i <= 13; i++)
                                {
                                    ArrGntAmt[i] = 0;
                                }
                            }
                        }

                        Create_Html2();
                        RetVal = Send_eMail();
                        if (RetVal.Trim() == "True")
                        {
                            cnt = cnt + 1;
                            oDS_PH_PY118B.SetValue("U_Col07", oRow, "Success");
                        }
                        else
                        {
                            oDS_PH_PY118B.SetValue("U_Col07", oRow, codeHelpClass.Left("Failure:" + RetVal, 50));
                        }
                        oMat1.SetLineData(oRow + 1);
                    }
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("선택된 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조건과 일치하는 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("보내는 사람 주소는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("SMTP Server는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Password는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    //MDC_Globals.Sbo_Application.StatusBar.SetText("Send_eMail Error: " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        #region Raise_FormItemEvent
        //		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			string sQry = null;
        //			int i = 0;
        //			SAPbouiCOM.ComboBox oCombo = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1

        //					if (pval.BeforeAction == true) {
        //						if (pval.ItemUID == "1") {
        //							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //							////추가및 업데이트시에
        //							//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //								////찾기모드일때 조건조회
        //								//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //								if (Execution_Process() == false) {
        //									BubbleEvent = false;
        //									return;
        //								} else {
        //									BubbleEvent = false;
        //									oMat1.AutoResizeColumns();
        //									return;
        //								}
        //							}
        //						} else if (pval.ItemUID == "CBtn1") {
        //							if (oForm.Items.Item("MSTCOD").Enabled == true) {
        //								oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
        //								BubbleEvent = false;
        //							}
        //						}
        //					} else if (pval.BeforeAction == false) {
        //						if (pval.ItemUID == "Btn1") {
        //							eMail_Process();
        //						} else if (pval.ItemUID == "Btn2" | pval.ItemUID == "Btn3") {
        //							PH_PY118_FormItemEnabled();
        //							/// 임시테이블에 저장하기
        //							if (pval.ItemUID == "Btn3") {
        //								SetSendMail();
        //							}
        //						}
        //					}
        //					break;

        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					break;
        //				//            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //				//            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
        //				//            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //				//             If pval.BeforeAction = True And pval.ColUid = "STRDAT" And pval.CharPressed = 9 Then
        //				//                If Trim$(oForm.Items(pval.ItemUID).Specific.Value) = "" Then
        //				//                    Call SBO_Application.StatusBar.SetText("조회시작일을 입력하여 주십시오.", bmt_Short, smt_Error)
        //				//                    BubbleEvent = False
        //				//                End If
        //				//             ElseIf pval.BeforeAction = True And pval.ColUid = "ENDDAT" And pval.CharPressed = 9 Then
        //				//                If Trim$(oForm.Items(pval.ItemUID).Specific.Value) = "" Then
        //				//                    Call SBO_Application.StatusBar.SetText("조회종료일을 입력하여 주십시오.", bmt_Short, smt_Error)
        //				//                    BubbleEvent = False
        //				//                End If
        //				//             End If
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					break;
        //				//            Select Case pval.ItemUid
        //				//            Case "Mat1", "Grid1"
        //				//                If pval.Row > 0 Then
        //				//                    oLastItemUID = pval.ItemUid
        //				//                    oLastColUID = pval.ColUID
        //				//                    oLastColRow = pval.Row
        //				//                End If
        //				//            Case Else
        //				//                oLastItemUID = pval.ItemUid
        //				//                oLastColUID = ""
        //				//                oLastColRow = 0
        //				//            End Select
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
        //							if (pval.ItemUID == "CLTCOD") {
        //								////기본사항 - 부서 (사업장에 따른 부서변경)
        //								oCombo = oForm.Items.Item("MSTDPT").Specific;

        //								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
        //								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
        //								sQry = sQry + " ORDER BY U_Code";

        //								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);

        //								if (oCombo.ValidValues.Count > 0) {
        //									oCombo.ValidValues.Add("%", "전체");
        //									oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //								}

        //								oForm.Items.Item("MSTDPT").DisplayDesc = true;
        //							}
        //						}
        //					}
        //					oForm.Freeze(false);
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == false & pval.ItemUID == "Mat1" & pval.ColUID == "Col11") {
        //						if (pval.Row == 0) {
        //							PrintChk();
        //						}
        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					if (pval.FormUID == oForm.UniqueID & pval.BeforeAction == false & pval.ItemUID == "Mat1" & pval.ColUID == "Col11") {
        //						if (pval.Row == 0) {
        //							PrintChk();
        //						}
        //					}
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
        //					if (pval.BeforeAction == false & pval.ItemChanged == true) {
        //						switch (pval.ItemUID) {
        //							case "MSTCOD":
        //								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.String)) {
        //									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oForm.Items.Item("MSTNAM").Specific.VALUE = "";
        //								} else {
        //									//UPGRADE_WARNING: oForm.Items(MSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oForm.Items.Item("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData(ref "U_FullName", ref "Code", ref "[@PH_PY001A]", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "");
        //								}
        //								oForm.Items.Item("MSTNAM").Update();
        //								break;

        //							case "STRDAT":
        //								//UPGRADE_WARNING: oForm.Items(STRDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("STRDAT").Specific.String)) {
        //									//UPGRADE_WARNING: oForm.Items(STRDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if (string.IsNullOrEmpty(oForm.Items.Item("STRDAT").Specific.String)) {
        //										//UPGRADE_WARNING: oForm.Items(STRDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										oForm.Items.Item("STRDAT").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
        //									}
        //								}
        //								break;

        //							case "ENDDAT":
        //								//UPGRADE_WARNING: oForm.Items(ENDDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("ENDDAT").Specific.String)) {
        //									//UPGRADE_WARNING: oForm.Items(ENDDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if (string.IsNullOrEmpty(oForm.Items.Item("ENDDAT").Specific.String)) {
        //										//UPGRADE_WARNING: oForm.Items(ENDDAT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										oForm.Items.Item("ENDDAT").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
        //									}
        //								}
        //								break;
        //						}
        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					break;
        //				//            If pval.BeforeAction = True Then
        //				//            ElseIf pval.BeforeAction = False Then
        //				//                oMat1.LoadFromDataSource
        //				//
        //				//                Call PH_PY118_FormItemEnabled
        //				//                Call PH_PY118_AddMatrixRow
        //				//
        //				//            End If
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
        //					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //					//컬렉션에서 삭제및 모든 메모리 제거
        //					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //					if (pval.BeforeAction == true) {
        //					} else if (pval.BeforeAction == false) {
        //						SubMain.RemoveForms(oFormUniqueID);
        //						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oForm = null;
        //						//UPGRADE_NOTE: oDS_PH_PY118B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oDS_PH_PY118B = null;

        //						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oMat1 = null;

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
        //					break;
        //				//            If pval.BeforeAction = True Then
        //				//
        //				//            ElseIf pval.BeforeAction = False Then
        //				//
        //				//            End If
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
        //					break;
        //				//            If pval.BeforeAction = True Then
        //				//
        //				//            ElseIf pval.Before_Action = False Then
        //				//
        //				//            End If
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
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormMenuEvent
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
        //				}
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1283":
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //						PH_PY118_FormItemEnabled();
        //						PH_PY118_AddMatrixRow();
        //						break;
        //					case "1284":
        //						break;
        //					case "1286":
        //						break;
        //					//            Case "1293":
        //					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
        //					case "1281":
        //						////문서찾기
        //						PH_PY118_FormItemEnabled();
        //						PH_PY118_AddMatrixRow();
        //						break;
        //					case "1282":
        //						////문서추가
        //						PH_PY118_FormItemEnabled();
        //						PH_PY118_AddMatrixRow();
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						PH_PY118_FormItemEnabled();
        //						break;
        //					case "1293":
        //						//// 행삭제
        //						break;
        //					//                '// [MAT1 용]
        //					//                 If oMat1.RowCount <> oMat1.VisualRowCount Then
        //					//                    oMat1.FlushToDataSource
        //					//
        //					//                    While (i <= oDS_PH_PY118B.Size - 1)
        //					//                        If oDS_PH_PY118B.GetValue("U_FILD01", i) = "" Then
        //					//                            oDS_PH_PY118B.RemoveRecord (i)
        //					//                            i = 0
        //					//                        Else
        //					//                            i = i + 1
        //					//                        End If
        //					//                    Wend
        //					//
        //					//                    For i = 0 To oDS_PH_PY118B.Size
        //					//                        Call oDS_PH_PY118B.setValue("U_LineNum", i, i + 1)
        //					//                    Next i
        //					//
        //					//                    oMat1.LoadFromDataSource
        //					//                End If
        //					//                Call PH_PY118_AddMatrixRow
        //				}
        //			}
        //			oForm.Freeze(false);
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if ((BusinessObjectInfo.BeforeAction == true)) {
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
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
        //			return;
        //			Raise_FormDataEvent_Error:


        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}
        //			switch (pval.ItemUID) {
        //				case "Mat1":
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
        #endregion







        #region 백업소스코드_S

        #region PH_PY118_AddMatrixRow
        //public void PH_PY118_AddMatrixRow()
        //{
        //    int oRow = 0;

        //    oForm.Freeze(true);

        //    //    '//[Mat1 용]
        //    //    oMat1.FlushToDataSource
        //    //    oRow = oMat1.VisualRowCount
        //    //
        //    //    If oMat1.VisualRowCount > 0 Then
        //    //        If Trim(oDS_PH_PY118B.GetValue("U_FILD01", oRow - 1)) <> "" Then
        //    //            If oDS_PH_PY118B.Size <= oMat1.VisualRowCount Then
        //    //                oDS_PH_PY118B.InsertRecord (oRow)
        //    //            End If
        //    //            oDS_PH_PY118B.Offset = oRow
        //    //            oDS_PH_PY118B.setValue "U_LineNum", oRow, oRow + 1
        //    //            oDS_PH_PY118B.setValue "U_FILD01", oRow, ""
        //    //            oDS_PH_PY118B.setValue "U_FILD02", oRow, ""
        //    //            oDS_PH_PY118B.setValue "U_FILD03", oRow, 0
        //    //            oMat1.LoadFromDataSource
        //    //        Else
        //    //            oDS_PH_PY118B.Offset = oRow - 1
        //    //            oDS_PH_PY118B.setValue "U_LineNum", oRow - 1, oRow
        //    //            oDS_PH_PY118B.setValue "U_FILD01", oRow - 1, ""
        //    //            oDS_PH_PY118B.setValue "U_FILD02", oRow - 1, ""
        //    //            oDS_PH_PY118B.setValue "U_FILD03", oRow - 1, 0
        //    //            oMat1.LoadFromDataSource
        //    //        End If
        //    //    ElseIf oMat1.VisualRowCount = 0 Then
        //    //        oDS_PH_PY118B.Offset = oRow
        //    //        oDS_PH_PY118B.setValue "U_LineNum", oRow, oRow + 1
        //    //        oDS_PH_PY118B.setValue "U_FILD01", oRow, ""
        //    //        oDS_PH_PY118B.setValue "U_FILD02", oRow, ""
        //    //        oDS_PH_PY118B.setValue "U_FILD03", oRow, 0
        //    //        oMat1.LoadFromDataSource
        //    //    End If

        //    oForm.Freeze(false);
        //    return;
        //PH_PY118_AddMatrixRow_Error:
        //    oForm.Freeze(false);
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY118_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PH_PY118_Print_Report01
        //		private void PH_PY118_Print_Report01()
        //		{

        //			string DocNum = null;
        //			short ErrNum = 0;
        //			string WinTitle = null;
        //			string ReportName = null;
        //			string sQry = null;

        //			string BPLID = null;
        //			string ItmBsort = null;
        //			string DocDate = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			/// ODBC 연결 체크
        //			if (ConnectODBC() == false) {
        //				goto PH_PY118_Print_Report01_Error;
        //			}

        //			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //			WinTitle = "[S142] 발주서";
        //			ReportName = "S142_1.rpt";
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			sQry = "EXEC PH_PY118_1 '" + oForm.Items.Item("8").Specific.VALUE + "'";
        //			MDC_Globals.gRpt_Formula = new string[2];
        //			MDC_Globals.gRpt_Formula_Value = new string[2];
        //			MDC_Globals.gRpt_SRptSqry = new string[2];
        //			MDC_Globals.gRpt_SRptName = new string[2];
        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //			/// Formula 수식필드

        //			/// SubReport


        //			/// Procedure 실행"
        //			sQry = "EXEC [PS_PP820_01] '" + BPLID + "', '" + ItmBsort + "', '" + DocDate + "'";

        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //					MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				}
        //			} else {
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("조회된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}

        //			return;
        //			PH_PY118_Print_Report01_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY118_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region PH_PY118_FormClear
        //		public void PH_PY118_FormClear()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string DocEntry = null;
        //			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY118'", ref "");
        //			if (Convert.ToDouble(DocEntry) == 0) {
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
        //			} else {
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //			}
        //			return;
        //			PH_PY118_FormClear_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY118_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region PH_PY118_DataValidCheck
        //		public bool PH_PY118_DataValidCheck()
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			functionReturnValue = false;
        //			int i = 0;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////----------------------------------------------------------------------------------
        //			////필수 체크
        //			////----------------------------------------------------------------------------------
        //			//    If Trim(oDS_PH_PY118A.GetValue("Code", 0)) = "" Then
        //			//        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
        //			//        oForm.Items("Code").CLICK ct_Regular
        //			//        Exit Function
        //			//    End If
        //			//
        //			//    oMat1.FlushToDataSource
        //			//    '// Matrix 마지막 행 삭제(DB 저장시)
        //			//    If oDS_PH_PY118B.Size > 1 Then oDS_PH_PY118B.RemoveRecord (oDS_PH_PY118B.Size - 1)
        //			//    oMat1.LoadFromDataSource

        //			functionReturnValue = true;
        //			return functionReturnValue;


        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			PH_PY118_DataValidCheck_Error:


        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = false;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY118_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}
        #endregion

        #region PH_PY118_Validate
        //		public bool PH_PY118_Validate(string ValidateType)
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
        //			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY118A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY118A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				goto PH_PY118_Validate_Exit;
        //			}
        //			//
        //			if (ValidateType == "수정") {

        //			} else if (ValidateType == "행삭제") {

        //			} else if (ValidateType == "취소") {

        //			}
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY118_Validate_Exit:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY118_Validate_Error:
        //			functionReturnValue = false;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY118_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}
        #endregion

        #endregion 백업소스코드_E

    }
}

