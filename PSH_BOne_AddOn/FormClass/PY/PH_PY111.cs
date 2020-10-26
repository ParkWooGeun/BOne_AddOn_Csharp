using System;
using System.Timers;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여 계산
    /// </summary>
    internal class PH_PY111 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        private SAPbouiCOM.DBDataSource oDS_PH_PY111A;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        #region 내부클래스 선언(구조체를 마이그레이션)
        /// <summary>
        /// 기초세액 내부 클래스
        /// </summary>
        internal class WG01CODR
        {
            internal double[] TB1AMT = new double[10];
            internal double[] TB1GON = new double[10];
            internal double[] TB1RAT = new double[10];
            internal double[] TB1KUM = new double[10];
        }

        WG01CODR WG01 = new WG01CODR(); //클래스 전체에서 사용됨, 클래스 레벨로 인스턴스 생성

        internal class WG03TILR
        {
            internal string[] CSUCOD = new string[24];
            internal string[] CSUNAM = new string[24];
            internal string[] MPYGBN = new string[24]; //월정급여
            internal double[] CSUKUM = new double[24]; //수당한도금액
            internal string[] GWATYP = new string[24]; //과세구분
            internal string[] GBHGBN = new string[24]; //고용보험여부
            internal string[] ROUNDT = new string[24]; //사사오입구분(끝전처리)
            internal short[] RODLEN = new short[24]; //끝전처리자릿수
            internal string[] GONSIL = new string[24]; //급여수식
            internal string[] BNSUSE = new string[24]; //상여항목
            internal string[] BTXCOD = new string[24]; //비과세코드
        }

        WG03TILR WK_C = new WG03TILR();

        internal class WG04TILR
        {
            internal string[] GONCOD = new string[19];
            internal string[] GONNAM = new string[19];
            internal string[] BNSUSE = new string[19];
            internal string[] GONSIL = new string[19];
            internal string[] ROUNDT = new string[19];
            internal short[] RODLEN = new short[19];
        }

        WG04TILR WK_G = new WG04TILR();
        #endregion

        #region 변수 선언
        private string oYM = string.Empty; //귀속연월
        private string oJOBTYP = string.Empty; //지급종류
        private string oJOBGBN = string.Empty; //지급구분
        private string oJIGBIL = string.Empty; //지급일자
        private string oJOBTRG = string.Empty; //지급대상구분
        private string oCLTCOD = string.Empty; //사업장
        private string oSTRDPT = string.Empty; //부서시작코드
        private string oENDDPT = string.Empty; //부서종료코드
        private string oMSTCOD = string.Empty; //사원번호
        private string oJSNCHK = string.Empty; //연말정산포함
        private string oYCHCHK = string.Empty; //연차지급포함

        private string oBNSCAL = string.Empty; //상여계산방법
        private string oBNSRAT = string.Empty; //상여율
        private string oSTRTAX = string.Empty; //세액대상기간시작(급여)
        private string oENDTAX = string.Empty; //세액대상기간종료(급여)
        private string oSTRBNS = string.Empty; //세액대상기간시작(상여)
        private string oENDBNS = string.Empty; //세액대상기간종료(상여)
        private string oGNEDAT = string.Empty; //상여계산기준일
        private string oEXPDAT = string.Empty; //퇴사자제외일
        private string oRETCHK = string.Empty; //상여퇴직임금에포함

        private bool G06_CHK; //소득세정산
        private bool G07_CHK; //주민세정산
        private bool G08_CHK; //건강보험정산
        private bool G90_CHK; //국민연금정산
        private bool G91_CHK; //농특세정산
        
        private string G04_BNSUSE;

        private string U_CSUCOD;
        private string U_GONCOD;
        #endregion


        //미사용 변수들_S
        //private short G_TotCnt = 0; //대상인원수
        //private short G_PayCnt = 0; //계산인원수
        //private short G_ChkCnt = 0; //잠금제외수

        //private string StrDate;
        //private string EndDate;
        //private int MaxRow;
        //private short JSNYER;

        //private bool G92_CHK; //고용보험정산

        //private string PAY_001;
        //private string PAY_007;

        //private double WK_TIMAMT; //시    급
        //private double WK_DAYAMT; //일    급
        //private double WK_STDAMT; //월    급
        //private double WK_BNSAMT; //상여기본
        //private double WK_APPBNS; //적용상여금
        //private double WK_AVRAMT; //급여기본등록의 평균임금

        //private string TB1_BT3COD; //사용하는 국외비과세코드
        //private string TB1_BT5COD; //사용하는 연구비과세코드

        //private double X01_Val; //X01
        //private double X02_Val; //X02
        //private double X03_Val; //X03
        //private double X04_Val; //X04

        //private short X10_Val;
        //private short X11_Val;
        //private short X12_Val;
        //private short X13_Val;
        //private string X14_Val;
        //private double X15_Val;
        //private short X16_Val;
        //private double X17_Val;
        //private short X18_Val;
        //private short X19_Val;
        //private short X20_Val;

        //private string REMARK1;
        //private string REMARK2;
        //private string REMARK3;
        //private bool TermCHK;
        //미사용 변수들_E

        private string tDocEntry = string.Empty; //저장전 문서번호 저장
        private bool CalcYN; //계산 여부

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            string strXml = string.Empty;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY111.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY111_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY111");

                strXml = oXmlDoc.xml.ToString();
                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY111_CreateItems();
                PH_PY111_EnableMenus();
                PH_PY111_SetDocument(oFormDocEntry01);
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
                oForm.ActiveItem = "CLTCOD";
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY111_CreateItems()
        {
            string sQry = string.Empty;
            short iCol = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            WG01CODR WG01 = new WG01CODR();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY111A = oForm.DataSources.DBDataSources.Item("@PH_PY111A");
                
                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //귀속연월
                oDS_PH_PY111A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
                oDS_PH_PY111A.SetValue("U_CHGCHK", 0, "N");
                oDS_PH_PY111A.SetValue("U_TAXCHK", 0, "Y");

                //소득세계산
                //지급종류
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").DisplayDesc = true;
                oForm.Items.Item("JOBTYP").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //지급구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;
                oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //급여대상자구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBTRG").Specific, "");
                oForm.Items.Item("JOBTRG").Specific.ValidValues.Add("%", "모두");
                oForm.Items.Item("JOBTRG").DisplayDesc = true;
                oForm.Items.Item("JOBTRG").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //직원구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P126' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JIGTYP").Specific, "");
                oForm.Items.Item("JIGTYP").Specific.ValidValues.Add("%", "모두");
                oForm.Items.Item("JIGTYP").DisplayDesc = true;
                oForm.Items.Item("JIGTYP").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //부서
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("STRDPT").Specific, "");
                oForm.Items.Item("STRDPT").Specific.ValidValues.Add("%", "모두");
                oForm.Items.Item("STRDPT").DisplayDesc = true;
                oForm.Items.Item("STRDPT").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ENDDPT").Specific, "");
                oForm.Items.Item("ENDDPT").Specific.ValidValues.Add("%", "모두");
                oForm.Items.Item("ENDDPT").DisplayDesc = true;
                oForm.Items.Item("ENDDPT").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //Check 버튼
                oForm.Items.Item("JSNCHK").Specific.ValOff = "N";
                oForm.Items.Item("JSNCHK").Specific.ValOn = "Y";

                oForm.Items.Item("YCHCHK").Specific.ValOff = "N";
                oForm.Items.Item("YCHCHK").Specific.ValOn = "Y";

                oForm.Items.Item("RETCHK").Specific.ValOff = "N";
                oForm.Items.Item("RETCHK").Specific.ValOn = "Y";
                oForm.Items.Item("RETCHK").Specific.Checked = true;

                oForm.Items.Item("TAXCHK").Specific.ValOff = "N";
                oForm.Items.Item("TAXCHK").Specific.ValOn = "Y";
                oForm.Items.Item("TAXCHK").Specific.Checked = true;

                oForm.Items.Item("GBHCHK").Specific.ValOff = "N";
                oForm.Items.Item("GBHCHK").Specific.ValOn = "Y";
                oForm.Items.Item("GBHCHK").Specific.Checked = true;

                oForm.Items.Item("CHGCHK").Specific.ValOff = "N";
                oForm.Items.Item("CHGCHK").Specific.ValOn = "Y";

                //상여
                for (iCol = 1; iCol <= 8; iCol++)
                {
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("1", "개월 이상");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("2", "개월 미만");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("3", "일수 이상");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("4", "일수 미만");
                    if (oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Count > 0)
                    {
                        oDS_PH_PY111A.SetValue("U_AP" + iCol + "GBN", 0, "1");
                    }
                    oForm.Items.Item("AP" + iCol + "GBN").DisplayDesc = true;
                }

                //상여계산방법
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P215' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("bBNSCAL").Specific, "");
                oForm.Items.Item("bBNSCAL").DisplayDesc = true;

                //91 기초세액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                for (iCol = 0; iCol < 10; iCol++)
                {
                    WG01.TB1AMT[iCol] = 0;
                    WG01.TB1GON[iCol] = 0;
                    WG01.TB1RAT[iCol] = 0;
                    WG01.TB1KUM[iCol] = 0;
                }

                sQry = "        SELECT      U_CODNBR,";
                sQry = sQry + "             U_CODAMT,";
                sQry = sQry + "             U_CODGON,";
                sQry = sQry + "             U_CODRAT,";
                sQry = sQry + "             U_CODKUM ";
                sQry = sQry + " FROM        [@PH_PY100B]";
                sQry = sQry + " WHERE       CODE = (";
                sQry = sQry + "                         SELECT      Top 1";
                sQry = sQry + "                                     Code";
                sQry = sQry + "                         FROM        [@PH_PY100A]";
                sQry = sQry + "                         WHERE       Code <= '" + DateTime.Now.ToString("yyyy") + "'";
                sQry = sQry + "                         ORDER BY    Code";
                sQry = sQry + "                                     Desc";
                sQry = sQry + "                     )";
                sQry = sQry + "             AND U_CODNBR BETWEEN '0001' AND '0010'";
                sQry = sQry + " ORDER BY    CODE,";
                sQry = sQry + "             U_CODNBR";
                sQry = sQry + "             DESC";
                oRecordSet.DoQuery(sQry);

                while (!(oRecordSet.EoF))
                {
                    WG01.TB1AMT[Convert.ToDouble(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODAMT").Value);
                    WG01.TB1GON[Convert.ToDouble(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODGON").Value);
                    WG01.TB1RAT[Convert.ToDouble(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODRAT").Value);
                    WG01.TB1KUM[Convert.ToDouble(oRecordSet.Fields.Item("U_CODNBR").Value)] = Convert.ToDouble(oRecordSet.Fields.Item("U_CODKUM").Value);
                    oRecordSet.MoveNext();
                }

                CalcYN = false;
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
        private void PH_PY111_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1287", true); //복제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 초기화(Set)
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY111_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY111_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY111_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
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
        private void PH_PY111_FormItemEnabled()
        {   
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (oForm.Visible == true)
                    {
                        oForm.ActiveItem = "CLTCOD";
                    }
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Btn1").Visible = true;
                    
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                    PH_PY111_FormClear();

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oCLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).ToString().Trim();
                    
                    oDS_PH_PY111A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM")); //귀속연월
                    oDS_PH_PY111A.SetValue("U_CHGCHK", 0, "N");
                    
                    oForm.Items.Item("JOBTYP").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue); //지급종류
                    oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue); //지급구분
                    oForm.Items.Item("JOBTRG").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue); //급여대상자구분
                    oForm.Items.Item("Btn1").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("Btn1").Visible = true;
                    oForm.ActiveItem = "DocEntry";

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oCLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).ToString().Trim();
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("Btn1").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.ActiveItem = "CLTCOD";
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Btn1").Visible = true;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oCLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).ToString().Trim();
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("Btn1").Enabled = true;
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
        /// 화면 클리어
        /// </summary>
        private void PH_PY111_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY111'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY111_DataValidCheck()
        {
            bool returnValue = false;
            short errNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oYM = oDS_PH_PY111A.GetValue("U_YM", 0).ToString().Trim();
                oJOBTYP = oDS_PH_PY111A.GetValue("U_JOBTYP", 0).ToString().Trim();
                oJOBGBN = oDS_PH_PY111A.GetValue("U_JOBGBN", 0).ToString().Trim();
                oCLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).ToString().Trim();

                oSTRDPT = oDS_PH_PY111A.GetValue("U_STRDPT", 0).ToString().Trim();
                oENDDPT = oDS_PH_PY111A.GetValue("U_ENDDPT", 0).ToString().Trim();

                if (oENDDPT == "%")
                {
                    oENDDPT = "ZZZZZZZZ";
                }

                oJIGBIL = oDS_PH_PY111A.GetValue("U_JIGBIL", 0).ToString().Trim();
                oMSTCOD = oDS_PH_PY111A.GetValue("U_MSTCOD", 0).ToString().Trim();
                oJSNCHK = oDS_PH_PY111A.GetValue("U_JSNCHK", 0).ToString().Trim();
                oYCHCHK = oDS_PH_PY111A.GetValue("U_YCHCHK", 0).ToString().Trim();
                //상여관련
                oBNSCAL = oDS_PH_PY111A.GetValue("U_bBNSCAL", 0).ToString().Trim();
                oBNSRAT = oDS_PH_PY111A.GetValue("U_bBNSRAT", 0).ToString().Trim();
                oSTRTAX = oDS_PH_PY111A.GetValue("U_bPAYSTR", 0).ToString().Trim();
                oENDTAX = oDS_PH_PY111A.GetValue("U_bPAYEND", 0).ToString().Trim();
                oSTRBNS = oDS_PH_PY111A.GetValue("U_bBNSSTR", 0).ToString().Trim();
                oENDBNS = oDS_PH_PY111A.GetValue("U_bBNSEND", 0).ToString().Trim();

                oRETCHK = oDS_PH_PY111A.GetValue("U_RETCHK", 0).ToString().Trim();
                oGNEDAT = oDS_PH_PY111A.GetValue("U_bGNEDAT", 0).ToString().Trim();

                //Check
                if (dataHelpClass.ChkYearMonth(oYM) == false)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oJOBTYP))
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oJOBGBN))
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oJIGBIL))
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_JOBTRG", 0)))
                {
                    errNum = 5;
                    throw new Exception();
                }

                oJOBTRG = oForm.Items.Item("JOBTRG").Specific.Selected.Value.ToString().Trim();

                //상여계산시 체크
                if (oJOBTYP != "1")
                {
                    if (Convert.ToDouble(oBNSRAT) == 0)
                    {
                        errNum = 6;
                        throw new Exception();
                    }
                    else if (oSTRTAX.Length != 6 || oENDTAX.Length != 6)
                    {
                        errNum = 7;
                        throw new Exception();
                    }
                    else if (Convert.ToInt32(oSTRTAX) > Convert.ToInt32(oENDTAX))
                    {
                        errNum = 8;
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_bGNEDAT", 0)))
                    {
                        errNum = 9;
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_bEXPDAT", 0)))
                    {
                        errNum = 10;
                        throw new Exception();
                    }
                }

                if (oJOBTYP == "1")
                {
                    if (oJOBGBN == "3")
                    {
                        //잔여급여계산시 근태기준일 입력
                        if (string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_GTDateFr", 0)))
                        {
                            errNum = 11;
                            throw new Exception();
                        }
                        if (string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_GTDateTo", 0)))
                        {
                            errNum = 12;
                            throw new Exception();
                        }
                    }
                }

                returnValue = true;
            }
            catch(Exception ex)
            {                
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속 연월을 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급 종류를 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급 구분을 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급일자는 필수입니다. 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급 대상구분은 필수입니다. 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여율은 필수입니다. 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("세액대상기간은 필수입니다. 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 8)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("세액대상기간이 올바르지 않습니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 9)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여계산기준일은 필수입니다. 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 10)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여지급제한일자는 필수입니다. 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 11)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("잔여급여 계산시 근태기준일(시작)은 필수입니다. 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 12)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("잔여급여 계산시 근태기준일(종료)은 필수입니다. 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                returnValue = false;
            }

            return returnValue;
        }

        /// <summary>
        /// Pay_Calc
        /// </summary>
        /// <returns></returns>
        private bool Pay_Calc()
        {
            bool returnValue = false;

            //short errNum = 0;

            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string YM = string.Empty;
            string JOBTYP = string.Empty;
            string JOBGBN = string.Empty;
            string JOBTRG = string.Empty;
            string JIGBIL = string.Empty;
            string MSTCOD = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).ToString().Trim();
                YM = oDS_PH_PY111A.GetValue("U_YM", 0).ToString().Trim();
                JOBTYP = oDS_PH_PY111A.GetValue("U_JOBTYP", 0).ToString().Trim();
                JOBGBN = oDS_PH_PY111A.GetValue("U_JOBGBN", 0).ToString().Trim();
                JOBTRG = oDS_PH_PY111A.GetValue("U_JOBTRG", 0).ToString().Trim();
                JIGBIL = oDS_PH_PY111A.GetValue("U_JIGBIL", 0).ToString().Trim();
                MSTCOD = oDS_PH_PY111A.GetValue("U_MSTCOD", 0).ToString().Trim();

                sQry = "Select Count(*) From [@PH_PY112A] ";
                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "' AND U_YM = '" + YM + "'";
                sQry = sQry + " AND U_JOBTYP = '" + JOBTYP + "'";
                sQry = sQry + " AND U_JOBGBN = '" + JOBGBN + "'";
                sQry = sQry + " AND U_JOBTRG = '" + JOBTRG + "'";
                sQry = sQry + " AND U_JIGBIL = '" + JIGBIL + "'";
                sQry = sQry + " AND U_MSTCOD LIKE '%" + MSTCOD + "%'";

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    if (PSH_Globals.SBO_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") == 2)
                    {
                        returnValue = false;
                        //errNum = 1;
                        //throw new Exception();
                    }
                    else
                    {
                        returnValue = true;
                    }
                }
                else
                {
                    returnValue = true;
                }
            }
            catch(Exception ex)
            {
                //if (errNum == 1) //메시지박스에서 "아니오"를 눌렀을 때
                //{
                //    //아무것도 하지 않고, 종료
                //}
                //else
                //{
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
                
                returnValue = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// Display_BonussRate
        /// </summary>
        private void Display_BonussRate()
        {
            string sQry = string.Empty;
            short iCol = 0;
            int iBNSMON = 0;
            string JOBGBN = string.Empty;

            //short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet2 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //소급일때 상여와 동일하게
                if (oJOBGBN == "2")
                {
                    JOBGBN = "1";
                }
                else
                {
                    JOBGBN = oJOBGBN;
                }

                sQry = " SELECT U_BNSCAL, U_BNSMON, U_BNSRAT, U_AP1MON, U_AP2MON, U_AP3MON, U_AP4MON, U_AP5MON,";
                sQry = sQry + " U_AP6MON, U_AP7MON, U_AP8MON, U_AP1RAT, U_AP2RAT, U_AP3RAT, U_AP4RAT, U_AP5RAT, ";
                sQry = sQry + " U_AP6RAT, U_AP7RAT, U_AP8RAT, U_AP1AMT, U_AP2AMT, U_AP3AMT, U_AP4AMT, U_AP5AMT,";
                sQry = sQry + " U_AP6AMT, U_AP7AMT, U_AP8AMT, U_AP1GBN, U_AP2GBN, U_AP3GBN, U_AP4GBN, U_AP5GBN,";
                sQry = sQry + " U_AP6GBN, U_AP7GBN, U_AP8GBN  FROM [@PH_PY108A] ";
                sQry = sQry + " WHERE U_CLTCOD = '" + oCLTCOD + "'  AND U_JOBGBN = '" + JOBGBN + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0 || oJOBTYP == "1")
                {
                    oDS_PH_PY111A.SetValue("U_bBNSCAL", 0, "1");
                    oDS_PH_PY111A.SetValue("U_bBNSRAT", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH1", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH2", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH3", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH4", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH5", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH6", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH7", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bMONTH8", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT1", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT2", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT3", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT4", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT5", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT6", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT7", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPRAT8", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT1", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT2", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT3", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT4", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT5", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT6", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT7", 0, "0");
                    oDS_PH_PY111A.SetValue("U_bAPPAMT8", 0, "0");

                    for (iCol = 1; iCol <= 8; iCol++)
                    {
                        oDS_PH_PY111A.SetValue("U_AP" + iCol + "GBN", 0, "1");
                    }

                    //2010.04.05 최동권 추가
                    oYM = oDS_PH_PY111A.GetValue("U_YM", 0).ToString().Trim();

                    oDS_PH_PY111A.SetValue("U_bPAYSTR", 0, oYM);
                    oDS_PH_PY111A.SetValue("U_bPAYEND", 0, oYM);
                    oDS_PH_PY111A.SetValue("U_bBNSSTR", 0, oYM);
                    oDS_PH_PY111A.SetValue("U_bBNSEND", 0, oYM);

                    oForm.Items.Item("bPAYSTR").Update();
                    oForm.Items.Item("bPAYEND").Update();
                    oForm.Items.Item("bBNSSTR").Update();
                    oForm.Items.Item("bBNSEND").Update();
                }
                else
                {
                    oDS_PH_PY111A.SetValue("U_bBNSCAL", 0, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY111A.SetValue("U_bBNSRAT", 0, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH1", 0, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH2", 0, oRecordSet.Fields.Item(4).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH3", 0, oRecordSet.Fields.Item(5).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH4", 0, oRecordSet.Fields.Item(6).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH5", 0, oRecordSet.Fields.Item(7).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH6", 0, oRecordSet.Fields.Item(8).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH7", 0, oRecordSet.Fields.Item(9).Value);
                    oDS_PH_PY111A.SetValue("U_bMONTH8", 0, oRecordSet.Fields.Item(10).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT1", 0, oRecordSet.Fields.Item(11).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT2", 0, oRecordSet.Fields.Item(12).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT3", 0, oRecordSet.Fields.Item(13).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT4", 0, oRecordSet.Fields.Item(14).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT5", 0, oRecordSet.Fields.Item(15).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT6", 0, oRecordSet.Fields.Item(16).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT7", 0, oRecordSet.Fields.Item(17).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPRAT8", 0, oRecordSet.Fields.Item(18).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT1", 0, oRecordSet.Fields.Item(19).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT2", 0, oRecordSet.Fields.Item(20).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT3", 0, oRecordSet.Fields.Item(21).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT4", 0, oRecordSet.Fields.Item(22).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT5", 0, oRecordSet.Fields.Item(23).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT6", 0, oRecordSet.Fields.Item(24).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT7", 0, oRecordSet.Fields.Item(25).Value);
                    oDS_PH_PY111A.SetValue("U_bAPPAMT8", 0, oRecordSet.Fields.Item(26).Value);

                    for (iCol = 1; iCol <= 8; iCol++)
                    {
                        oDS_PH_PY111A.SetValue("U_AP" + iCol + "GBN", 0, oRecordSet.Fields.Item(26 + iCol).Value);
                    }

                    //2010.04.05 최동권 추가
                    oYM = oDS_PH_PY111A.GetValue("U_YM", 0).Trim();

                    if (!string.IsNullOrEmpty(oYM))
                    {
                        iBNSMON = oRecordSet.Fields.Item(1).Value * -1;
                        if (iBNSMON != 0)
                        {
                            iBNSMON = iBNSMON + 1;
                        }

                        oRecordSet2.DoQuery(("SELECT CONVERT(VARCHAR(6),DATEADD(MM, " + iBNSMON.ToString() + ", '" + oYM + "01'),112) FROM OADM"));

                        oDS_PH_PY111A.SetValue("U_bPAYSTR", 0, oRecordSet2.Fields.Item(0).Value);
                        oDS_PH_PY111A.SetValue("U_bPAYEND", 0, oYM);
                        oDS_PH_PY111A.SetValue("U_bBNSSTR", 0, oRecordSet2.Fields.Item(0).Value);
                        oDS_PH_PY111A.SetValue("U_bBNSEND", 0, oYM);

                        oForm.Items.Item("bPAYSTR").Update();
                        oForm.Items.Item("bPAYEND").Update();
                        oForm.Items.Item("bBNSSTR").Update();
                        oForm.Items.Item("bBNSEND").Update();
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //Create_Tiltle() = "Display_BonussRate :" + Strings.Space(10) + Err().Description;
            }
            finally
            {
                oForm.Update();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet2);
            }
        }

        /// <summary>
        /// Create_Tiltle
        /// </summary>
        /// <returns></returns>
        private string Create_Tiltle()
        {
            string returnValue = string.Empty;
            string sQry = string.Empty;

            short errNum = 0;

            int CSUCNT = 0;
            int GONCNT = 0;
            int iCol = 0;
            
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //1.항목명 초기화
                CSUCNT = 0;
                GONCNT = 0;
                for (iCol = 0; iCol < 24; iCol++)
                {
                    WK_C.CSUCOD[iCol] = "";
                    WK_C.CSUNAM[iCol] = "---";
                    WK_C.CSUKUM[iCol] = 0;
                    WK_C.MPYGBN[iCol] = "N";
                    WK_C.GWATYP[iCol] = "";
                    WK_C.GBHGBN[iCol] = "";
                    WK_C.ROUNDT[iCol] = "R";
                    WK_C.RODLEN[iCol] = 1;
                    WK_C.GONSIL[iCol] = "";
                    WK_C.BNSUSE[iCol] = "N";
                    WK_C.BTXCOD[iCol] = "";
                }

                for (iCol = 0; iCol < 18; iCol++)
                {
                    WK_G.GONCOD[iCol] = "";
                    WK_G.GONNAM[iCol] = "---";
                    WK_G.BNSUSE[iCol] = "N";
                    WK_G.GONSIL[iCol] = "";
                    WK_G.ROUNDT[iCol] = "R";
                    WK_G.RODLEN[iCol] = 1;
                }

                //2.수당항목 가져오기
                sQry = "Exec PH_PY102  '" + oCLTCOD + "' , '" + oYM + "', '', '', '', ''";
                oRecordSet.DoQuery(sQry);
                CSUCNT = oRecordSet.RecordCount;

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                    //수당항목관리를 확인하세요.
                    //goto Error_Message;
                }
                else
                {
                    iCol = 0;
                    while (!(oRecordSet.EoF))
                    {
                        WK_C.CSUCOD[iCol] = oRecordSet.Fields.Item("U_CSUCOD").Value;
                        WK_C.CSUNAM[iCol] = oRecordSet.Fields.Item("U_CSUNAM").Value;
                        WK_C.CSUKUM[iCol] = oRecordSet.Fields.Item("U_KUMAMT").Value;
                        WK_C.MPYGBN[iCol] = oRecordSet.Fields.Item("U_MONPAY").Value;
                        WK_C.GWATYP[iCol] = oRecordSet.Fields.Item("U_GWATYP").Value;
                        WK_C.GBHGBN[iCol] = oRecordSet.Fields.Item("U_GBHGBN").Value;
                        WK_C.ROUNDT[iCol] = oRecordSet.Fields.Item("U_ROUNDT").Value;
                        WK_C.RODLEN[iCol] = Convert.ToInt16(oRecordSet.Fields.Item("U_LENGTH").Value);
                        WK_C.BNSUSE[iCol] = oRecordSet.Fields.Item("U_BNSUSE").Value;
                        WK_C.BTXCOD[iCol] = oRecordSet.Fields.Item("U_BTXCOD").Value;
                        U_CSUCOD = oRecordSet.Fields.Item("Code").Value;
                        iCol = iCol + 1;
                        oRecordSet.MoveNext();
                    }
                }

                //3.공제항목 가져오기
                sQry = "Exec PH_PY103  '" + oCLTCOD + "' , '" + oYM + "',  '', ''";
                oRecordSet.DoQuery(sQry);
                GONCNT = oRecordSet.RecordCount;

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                    //공제항목관리를 확인하세요.
                    //goto Error_Message;
                }
                else
                {
                    iCol = 0;
                    G06_CHK = false; //소득세정산
                    G07_CHK = false; //주민세정산
                    G08_CHK = false; //건강보험정산
                    G90_CHK = false; //국민연금정산
                    G91_CHK = false; //농특세정산

                    while (!(oRecordSet.EoF))
                    {
                        WK_G.GONCOD[iCol] = oRecordSet.Fields.Item("U_CSUCOD").Value;
                        WK_G.GONNAM[iCol] = oRecordSet.Fields.Item("U_CSUNAM").Value;
                        WK_G.BNSUSE[iCol] = oRecordSet.Fields.Item("U_BNSUSE").Value; //상여
                        WK_G.GONSIL[iCol] = oRecordSet.Fields.Item("U_SILCUN").Value; //급여계산식
                        WK_G.ROUNDT[iCol] = oRecordSet.Fields.Item("U_ROUNDT").Value; //끝전처리
                        WK_G.RODLEN[iCol] = Convert.ToInt16(oRecordSet.Fields.Item("U_LENGTH").Value);

                        if (WK_G.GONCOD[iCol].Trim() == "G06")
                        {
                            G06_CHK = true;
                        }

                        if (WK_G.GONCOD[iCol].Trim() == "G07")
                        {
                            G07_CHK = true;
                        }
                            
                        if (WK_G.GONCOD[iCol].Trim() == "G08")
                        {
                            G08_CHK = true;
                        }
                            
                        if (WK_G.GONCOD[iCol].Trim() == "G90")
                        {
                            G90_CHK = true;
                        }
                            
                        if (WK_G.GONCOD[iCol].Trim() == "G91")
                        {
                            G91_CHK = true;
                        }
                            
                        if (WK_G.GONCOD[iCol].Trim() == "G04")
                        {
                            G04_BNSUSE = oRecordSet.Fields.Item("U_BNSUSE").Value;
                        }
                            
                        U_GONCOD = oRecordSet.Fields.Item("Code").Value;
                        iCol = iCol + 1;
                        oRecordSet.MoveNext();
                    }
                }

                //4.수당타이틀저장
                sQry = "SELECT CODE FROM [@PH_PY111T] WHERE U_YM = '" + oYM + "'";
                sQry = sQry + " AND U_JOBTYP = '" + oJOBTYP + "' AND U_JOBGBN = '" + oJOBGBN + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    //1.1) Insert
                    sQry = "INSERT INTO [@PH_PY111T]  (Code, Name, U_YM, U_JOBTYP, U_JOBGBN,";
                    sQry = sQry + " U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08,";
                    sQry = sQry + " U_CSUD09, U_CSUD10, U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16,";
                    sQry = sQry + " U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20, U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24,";
                    sQry = sQry + " U_GONG01, U_GONG02, U_GONG03, U_GONG04, U_GONG05, U_GONG06, U_GONG07, U_GONG08,";
                    sQry = sQry + " U_GONG09, U_GONG10, U_GONG11, U_GONG12, U_GONG13, U_GONG14, U_GONG15, U_GONG16,";
                    sQry = sQry + " U_GONG17, U_GONG18,";
                    sQry = sQry + " U_CSUC01, U_CSUC02, U_CSUC03, U_CSUC04, U_CSUC05, U_CSUC06, U_CSUC07, U_CSUC08,";
                    sQry = sQry + " U_CSUC09, U_CSUC10, U_CSUC11, U_CSUC12, U_CSUC13, U_CSUC14, U_CSUC15, U_CSUC16,";
                    sQry = sQry + " U_CSUC17, U_CSUC18, U_CSUC19, U_CSUC20, U_CSUC21, U_CSUC22, U_CSUC23, U_CSUC24,";
                    sQry = sQry + " U_GONC01, U_GONC02, U_GONC03, U_GONC04, U_GONC05, U_GONC06, U_GONC07, U_GONC08,";
                    sQry = sQry + " U_GONC09, U_GONC10, U_GONC11, U_GONC12, U_GONC13, U_GONC14, U_GONC15, U_GONC16,";
                    sQry = sQry + " U_GONC17, U_GONC18) VALUES ( ";
                    sQry = sQry + "  '" + oYM + oJOBTYP + oJOBGBN + "'";
                    sQry = sQry + ", '" + oYM + oJOBTYP + oJOBGBN + "'";
                    sQry = sQry + ", '" + oYM + " ', '" + oJOBTYP + "', '" + oJOBGBN + "'";

                    for (iCol = 0; iCol < 24; iCol++)
                    {
                        sQry = sQry + ", N'" + WK_C.CSUNAM[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 18; iCol++)
                    {
                        sQry = sQry + ", N'" + WK_G.GONNAM[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 24; iCol++)
                    {
                        sQry = sQry + ", N'" + WK_C.CSUCOD[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 18; iCol++)
                    {
                        sQry = sQry + ", N'" + WK_G.GONCOD[iCol] + "'";
                    }
                    sQry = sQry + ")";

                    oRecordSet.DoQuery(sQry);
                }
                else
                {
                    //1.2) Update
                    sQry = "UPDATE [@PH_PY111T] SET U_JOBGBN = '" + oJOBGBN + "'";

                    for (iCol = 0; iCol < 24; iCol++)
                    {
                        sQry = sQry + ", U_CSUD" + iCol.ToString().PadLeft(2, '0') + " = N'" + WK_C.CSUNAM[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 18; iCol++)
                    {
                        sQry = sQry + ", U_GONG" + iCol.ToString().PadLeft(2, '0') + " = N'" + WK_G.GONNAM[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 24; iCol++)
                    {
                        sQry = sQry + ", U_CSUC" + iCol.ToString().PadLeft(2, '0') + " = N'" + WK_C.CSUCOD[iCol] + "'";
                    }
                    for (iCol = 0; iCol < 18; iCol++)
                    {
                        sQry = sQry + ", U_GONC" + iCol.ToString().PadLeft(2, '0') + " = N'" + WK_G.GONCOD[iCol] + "'";
                    }
                    sQry = sQry + " WHERE   U_YM = '" + oYM + "'";
                    sQry = sQry + " AND     U_JOBTYP = '" + oJOBTYP + "'";
                    sQry = sQry + " AND     U_JOBGBN = '" + oJOBGBN + "'";

                    oRecordSet.DoQuery(sQry);
                }

                returnValue = "";
            }
            catch(Exception ex)
            {
                string stringSpace = new string(' ', 10);

                if (errNum == 1)
                {
                    returnValue = "(PH_PY102)수당항목자료가 없습니다. 확인하여 주십시오. ";
                }
                else if (errNum == 2)
                {
                    returnValue = "(PH_PY103)공제항목자료가 없습니다. 확인하여 주십시오. ";
                }
                else
                {
                    returnValue = "Create_Title : " + stringSpace + ex.Message;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// AddOn 연결 유지용 Timer 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void KeepAddOnConnection(object sender, ElapsedEventArgs e)
        {
            PSH_Globals.SBO_Application.RemoveWindowsMessage(BoWindowsMessageType.bo_WM_TIMER, true);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
           
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {                        tDocEntry = oForm.Items.Item("DocEntry").Specific.Value;
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY111_DataValidCheck() == false) //유효성 검사
                            {
                                BubbleEvent = false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CBtn1" && oForm.Items.Item("MSTCOD").Enabled == true) //ChooseBtn사원리스트
                    {
                        oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                    }
                    else if (pVal.ItemUID == "Btn1") //급(상)여계산
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                            PSH_Globals.SBO_Application.MessageBox("확인모드일 경우에만 급상여계산이 가능합니다.");
                        }
                        else
                        {
                            CalcYN = true;

                            tDocEntry = oForm.Items.Item("DocEntry").Specific.Value;

                            if (PH_PY111_DataValidCheck() == true && Pay_Calc() == true) //유효성 검사
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    Timer timer = new Timer();
                                    SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 100, false);

                                    timer.Interval = 30000; //30초
                                    timer.Elapsed += KeepAddOnConnection;
                                    timer.Start();

                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim() != "2")
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111 '" + tDocEntry + "'"); //정상계산
                                    }
                                    else
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111_SOGUB '" + tDocEntry + "'"); //소급계산
                                    }

                                    timer.Stop();
                                    timer.Dispose();

                                    ProgBar01.Stop();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);

                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                            }
                            else
                            {
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            if (CalcYN == true)
                            {
                                CalcYN = false;
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    Timer timer = new Timer();
                                    SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 100, false);

                                    timer.Interval = 30000; //30초
                                    timer.Elapsed += KeepAddOnConnection;
                                    timer.Start();
                                    
                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim() != "2")
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111 '" + tDocEntry + "'"); //정상계산
                                    }
                                    else
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111_SOGUB '" + tDocEntry + "'"); //소급계산
                                    }

                                    timer.Stop();
                                    timer.Dispose();

                                    ProgBar01.Stop();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);

                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PH_PY111_FormItemEnabled();

                                    oForm.Items.Item("DocEntry").Specific.Value = tDocEntry;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                            else if (CalcYN == false)
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PH_PY111_FormItemEnabled();

                                    oForm.Items.Item("DocEntry").Specific.Value = tDocEntry;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else
                                {
                                    PH_PY111_FormItemEnabled();
                                }
                            }
                        }
                    }

                    //수당/공제 계산식 반영제외
                    if (pVal.ItemUID == "CHGCHK")
                    {
                        if (oDS_PH_PY111A.GetValue("U_CHGCHK", 0).Trim() == "Y")
                        {
                            oDS_PH_PY111A.SetValue("U_TAXCHK", 0, "N"); //소득세 계산함
                            oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "N"); //고용보험 계산함
                        }
                    }
                    if (pVal.ItemUID == "TAZCHK" || pVal.ItemUID == "GBHCHK")
                    {
                        if (oDS_PH_PY111A.GetValue("U_TAXCHK", 0).Trim() == "Y" || oDS_PH_PY111A.GetValue("U_GBHCHK", 0).Trim() == "Y")
                        {
                            oDS_PH_PY111A.SetValue("U_CHGCHK", 0, "N");
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
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "YM")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim()))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속연월은 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                        }

                        if (pVal.ItemUID == "MSTCOD")
                        {
                            if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "") == true)
                            {
                                oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
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
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string JIGBIL = string.Empty;
            string queryString = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

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
                        if (pVal.ItemUID == "CLTCOD") //사업장
                        {
                            if (oForm.Items.Item("CLTCOD").Specific.Selected != null)
                            {
                                oCLTCOD = oDS_PH_PY111A.GetValue("U_CLTCOD", 0).Trim();
                                if (!string.IsNullOrEmpty(oCLTCOD))
                                {
                                    Display_BonussRate();
                                }
                            }
                            else
                            {
                                oJOBTYP = "";
                            }
                        }

                        if (pVal.ItemUID == "JOBTYP") //지급종류
                        {
                            if (oForm.Items.Item("JOBTYP").Specific.Selected != null)
                            {
                                oJOBTYP = oForm.Items.Item("JOBTYP").Specific.Selected.Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(oJOBTYP))
                                {
                                    Display_BonussRate();
                                }
                            }
                            else
                            {
                                oJOBTYP = "";
                            }

                            if (Convert.ToInt32(oYM) <= 201012)
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "Y");
                            }
                            else if ((oJOBTYP == "1" || oJOBTYP == "3") && oJOBGBN == "1")
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "Y");
                            }
                            else
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "N");
                            }
                            oForm.Items.Item("GBHCHK").Update();
                        }

                        if (pVal.ItemUID == "JOBGBN") //지급구분
                        {
                            if (oForm.Items.Item("JOBGBN").Specific.Selected != null)
                            {
                                oJOBGBN = oDS_PH_PY111A.GetValue("U_JOBGBN", 0).Trim();

                            }
                            else
                            {
                                oJOBGBN = "";
                            }
                            if (!string.IsNullOrEmpty(oJOBGBN))
                            {
                                Display_BonussRate();
                            }
                            if (Convert.ToInt32(oYM) <= 201012)
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "Y");
                            }
                            else if ((oJOBTYP == "1" || oJOBTYP == "3") && oJOBGBN == "1")
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "Y");
                            }
                            else
                            {
                                oDS_PH_PY111A.SetValue("U_GBHCHK", 0, "N");
                            }
                            oForm.Items.Item("GBHCHK").Update();
                        }

                        if (pVal.ItemUID == "JOBTRG") //급여지급대상일
                        {
                            if (oForm.Items.Item("JOBTRG").Specific.Selected != null)
                            {
                                oJOBTRG = oDS_PH_PY111A.GetValue("U_JOBTRG", 0).Trim();
                                if (!string.IsNullOrEmpty(oDS_PH_PY111A.GetValue("U_YM", 0).Trim()))
                                {
                                    oYM = oDS_PH_PY111A.GetValue("U_YM", 0).Trim();
                                    queryString = " SELECT DBO.Func_PAYTerm( '" + oYM.Substring(0, 4) + "-" + oYM.Substring(4, 2) + "-01" + "', '" + (oJOBTRG == "%" ? "1" : oJOBTRG) + "')";
                                    oRecordSet.DoQuery(queryString);
                                    if (oRecordSet.RecordCount > 0)
                                    {
                                        JIGBIL = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 18, 8);
                                        oGNEDAT = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 9, 8);
                                        oEXPDAT = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 27, 8);

                                        oDS_PH_PY111A.SetValue("U_JIGBIL", 0, JIGBIL);
                                        oDS_PH_PY111A.SetValue("U_bGNEDAT", 0, oGNEDAT);
                                        oDS_PH_PY111A.SetValue("U_bEXPDAT", 0, oEXPDAT);

                                        oForm.Items.Item("JIGBIL").Update();
                                        oForm.Items.Item("bGNEDAT").Update();
                                        oForm.Items.Item("bEXPDAT").Update();
                                    }
                                }
                            }
                            else
                            {
                                oJOBTRG = "";
                            }
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string JIGBIL = string.Empty;
            string queryString = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "YM") //귀속년월
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
                            {
                                oDS_PH_PY111A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM")); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM"));
                            }
                            else
                            {
                                oDS_PH_PY111A.SetValue("U_YM", 0, oForm.Items.Item("YM").Specific.Value);
                            }
                            oYM = oDS_PH_PY111A.GetValue("U_YM", 0).Trim();

                            if (!string.IsNullOrEmpty(oJOBTRG))
                            {
                                queryString = " SELECT DBO.Func_PAYTerm( '" + oYM.Substring(0, 4) + "-" + oYM.Substring(4, 2) + "-01" + "', '" + (oJOBTRG == "%" ? "1" : oJOBTRG) + "')";
                                oRecordSet.DoQuery(queryString);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    JIGBIL = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 18, 8);
                                    oGNEDAT = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 9, 8);
                                    oEXPDAT = codeHelpClass.Mid(oRecordSet.Fields.Item(0).Value, 27, 8);

                                    oDS_PH_PY111A.SetValue("U_JIGBIL", 0, JIGBIL);
                                    oDS_PH_PY111A.SetValue("U_bGNEDAT", 0, oGNEDAT);
                                    oDS_PH_PY111A.SetValue("U_bEXPDAT", 0, oEXPDAT);
                                    
                                    oForm.Items.Item("JIGBIL").Update();
                                }
                            }

                            if (!string.IsNullOrEmpty(oYM))
                            {
                                Display_BonussRate();
                            }
                        }

                        if (pVal.ItemUID == "MSTCOD") //사번
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                            {
                                oDS_PH_PY111A.SetValue("U_MSTCOD", 0, oForm.Items.Item("MSTCOD").Specific.Value);
                                oDS_PH_PY111A.SetValue("U_MSTNAM", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.String + "'", ""));
                            }
                            else
                            {
                                oDS_PH_PY111A.SetValue("U_MSTNAM", 0, "");
                            }
                            oForm.Items.Item("MSTNAM").Update();
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY111A);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY111A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY111_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
                        case "1281": //문서찾기
                            PH_PY111_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY111_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY111_FormItemEnabled();
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
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

                switch (pVal.ItemUID)
                {
                    case "Mat01":
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
