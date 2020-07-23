//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;

using System;

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

        //private SAPbobsCOM.Recordset sRecordset;
        #endregion

        #region 구조체 선언(내부클래스 마이그레이션 대상, 사용안됨)
        //private struct WG04TILR
        //{
        //    [VBFixedArray(18)]
        //    public string[] GONCOD;

        //    [VBFixedArray(18)]
        //    public string[] GONNAM;

        //    [VBFixedArray(18)]
        //    public string[] BNSUSE; /// 상여

        //    [VBFixedArray(18)]
        //    public string[] GONSIL; /// 계산식

        //    [VBFixedArray(18)]
        //    public string[] ROUNDT;

        //    [VBFixedArray(18)]
        //    public short[] RODLEN;

        //    public void Initialize()
        //    {
        //        GONCOD = new string[19];
        //        GONNAM = new string[19];
        //        BNSUSE = new string[19];
        //        GONSIL = new string[19];
        //        ROUNDT = new string[19];
        //        RODLEN = new short[19];
        //    }
        //}
        
        //WG04TILR WK_G;

        //private struct WG33PAYR
        //{
        //    public short DocNum; //문서번호

        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
        //    public char[] U_MSTCOD;

        //    [VBFixedString(50), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 50)]
        //    public char[] U_MSTNAM;

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_EmpID;

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_MSTBRK;

        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
        //    public char[] U_CLTCOD;

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_MSTDPT; //부서

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_MSTSTP; //직책

        //    [VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
        //    public char[] U_CLTNAM;

        //    [VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
        //    public char[] U_BRKNAM;

        //    [VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
        //    public char[] U_DPTNAM; //부서

        //    [VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
        //    public char[] U_STPNAM; //직책

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_PAYTYP;

        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
        //    public char[] U_JOBTRG; //급여지급일구분

        //    public double U_MEDAMT;
        //    public double U_KUKAMT;
        //    public double U_GBHAMT;
        //    public string U_PERNBR;

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_JIGCOD;

        //    [VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
        //    public char[] U_HOBONG;
                        
        //    public double U_STDAMT; //기 본 급
        //    public double U_BASAMT; //통상일급
        //    public double U_DAYAMT; //기본일급

        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
        //    public char[] U_INPDAT; //입사일자

        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
        //    public char[] U_INEDAT; //수습만료일
        //    [VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]

        //    public char[] U_OUTDAT; //퇴사일자
        //    public short U_TAXCNT;
        //    public short U_CHLCNT;
        //    public string U_NJCGBN;
        //    public double U_MEDFRG;
        //    [VBFixedArray(24)]
        //    public double[] U_CSUAMT;
        //    public double U_GWASEE;
        //    public double U_BTAX01;
        //    public double U_BTAX02;
        //    public double U_BTAX03;
        //    public double U_BTAX04;
        //    public double U_BTAX05;
        //    public double U_BTAX06;
        //    public double U_BTAX07;
        //    public double U_BTXG01;
        //    public double U_BTXH01;
        //    public double U_BTXH05;
        //    public double U_BTXH06;
        //    public double U_BTXH07;
        //    public double U_BTXH08;
        //    public double U_BTXH09;
        //    public double U_BTXH10;
        //    public double U_BTXH11;
        //    public double U_BTXH12;
        //    public double U_BTXH13;
        //    public double U_BTXI01;
        //    public double U_BTXK01;
        //    public double U_BTXM01;
        //    public double U_BTXM02;
        //    public double U_BTXM03;
        //    public double U_BTXO01;
        //    public double U_BTXQ01;
        //    public double U_BTXR10;
        //    public double U_BTXS01;
        //    public double U_BTXT01;
        //    public double U_BTXY01;
        //    public double U_BTXY02;
        //    public double U_BTXY03;
        //    public double U_BTXY21;
        //    public double U_BTXZ01;
        //    public double U_BTXY22;
        //    public double U_BTXX01;
        //    public double U_BTXY20;
        //    public double U_BTXTOT;
        //    public double U_TOTPAY;
        //    [VBFixedArray(18)]
        //    public double[] U_GONAMT;
        //    public double U_TOTGON;
        //    public double U_SILJIG;
        //    public double U_AVRPAY; //상여금
        //    public double U_NABTAX;
        //    public double U_BNSRAT;
        //    public double U_APPRAT;
        //    public short U_GNSYER;
        //    public short U_GNSMON;
        //    public short U_TAXTRM;
        //    public double U_BONUSS;

        //    public void Initialize()
        //    {
        //        U_CSUAMT = new double[25];
        //        U_GONAMT = new double[19];
        //    }
        //}

        //WG33PAYR WG03;

        #region WG01CODR
        //private struct WG01CODR
        //{
        //    [VBFixedArray(10)]
        //    public double[] TB1AMT;

        //    [VBFixedArray(10)]
        //    public double[] TB1GON;

        //    [VBFixedArray(10)]
        //    public double[] TB1RAT;

        //    [VBFixedArray(10)]
        //    public double[] TB1KUM;

        //    public void Initialize()
        //    {
        //        TB1AMT = new double[11];
        //        TB1GON = new double[11];
        //        TB1RAT = new double[11];
        //        TB1KUM = new double[11];
        //    }
        //}

        //WG01CODR WG01;
        #endregion

        #region WG03TILR
        //private struct WG03TILR
        //{
        //    [VBFixedArray(24)]
        //    public string[] CSUCOD;

        //    [VBFixedArray(24)]
        //    public string[] CSUNAM;

        //    [VBFixedArray(24)]
        //    public string[] MPYGBN; //월정급여

        //    [VBFixedArray(24)]
        //    public double[] CSUKUM; //수당한도금액

        //    [VBFixedArray(24)]
        //    public string[] GWATYP; //과세구분

        //    [VBFixedArray(24)]
        //    public string[] GBHGBN; //고용보험여부

        //    [VBFixedArray(24)]
        //    public string[] ROUNDT; //사사오입구분(끝전처리)

        //    [VBFixedArray(24)]
        //    public short[] RODLEN; //끝전처리자릿수

        //    [VBFixedArray(24)]
        //    public string[] GONSIL; //급여수식

        //    [VBFixedArray(24)]
        //    public string[] BNSUSE; //상여항목

        //    [VBFixedArray(24)]            
        //    public string[] BTXCOD; //비과세코드

        //    //UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        //    public void Initialize()
        //    {
        //        //UPGRADE_WARNING: CSUCOD 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        //        CSUCOD = new string[25];
        //        CSUNAM = new string[25];
        //        MPYGBN = new string[25];
        //        CSUKUM = new double[25];
        //        GWATYP = new string[25];
        //        GBHGBN = new string[25];
        //        ROUNDT = new string[25];
        //        RODLEN = new short[25];
        //        GONSIL = new string[25];
        //        BNSUSE = new string[25];
        //        BTXCOD = new string[25];
        //    }
        //}
        ////UPGRADE_WARNING: WK_C 구조체의 배열은 사용하기 전에 초기화해야 합니다.
        //WG03TILR WK_C;
        #endregion

        #endregion

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
                PH_PY111_SetDocument(oFromDocEntry01);
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
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY111_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY111_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY111_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
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
                    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim() != "2")
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111 '" + tDocEntry + "'"); //정상계산
                                    }
                                    else
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111_SOGUB '" + tDocEntry + "'"); //소급계산
                                    }

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
                                    PSH_Globals.SBO_Application.StatusBar.SetText("급상여계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim() != "2")
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111 '" + tDocEntry + "'"); //정상계산
                                    }
                                    else
                                    {
                                        oRecordSet.DoQuery("EXEC PH_PY111_SOGUB '" + tDocEntry + "'"); //소급계산
                                    }

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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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

        #region 백업소스코드_S

        #region PH_PY111_AddMatrixRow
        //		public void PH_PY111_AddMatrixRow()
        //		{
        //			int oRow = 0;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			//    '//[Mat1 용]
        //			//    oMat1.FlushToDataSource
        //			//    oRow = oMat1.VisualRowCount
        //			//
        //			//    If oMat1.VisualRowCount > 0 Then
        //			//        If Trim(oDS_PH_PY111B.GetValue("U_FILD01", oRow - 1)) <> "" Then
        //			//            If oDS_PH_PY111B.Size <= oMat1.VisualRowCount Then
        //			//                oDS_PH_PY111B.InsertRecord (oRow)
        //			//            End If
        //			//            oDS_PH_PY111B.Offset = oRow
        //			//            oDS_PH_PY111B.setValue "U_LineNum", oRow, oRow + 1
        //			//            oDS_PH_PY111B.setValue "U_FILD01", oRow, ""
        //			//            oDS_PH_PY111B.setValue "U_FILD02", oRow, ""
        //			//            oDS_PH_PY111B.setValue "U_FILD03", oRow, 0
        //			//            oMat1.LoadFromDataSource
        //			//        Else
        //			//            oDS_PH_PY111B.Offset = oRow - 1
        //			//            oDS_PH_PY111B.setValue "U_LineNum", oRow - 1, oRow
        //			//            oDS_PH_PY111B.setValue "U_FILD01", oRow - 1, ""
        //			//            oDS_PH_PY111B.setValue "U_FILD02", oRow - 1, ""
        //			//            oDS_PH_PY111B.setValue "U_FILD03", oRow - 1, 0
        //			//            oMat1.LoadFromDataSource
        //			//        End If
        //			//    ElseIf oMat1.VisualRowCount = 0 Then
        //			//        oDS_PH_PY111B.Offset = oRow
        //			//        oDS_PH_PY111B.setValue "U_LineNum", oRow, oRow + 1
        //			//        oDS_PH_PY111B.setValue "U_FILD01", oRow, ""
        //			//        oDS_PH_PY111B.setValue "U_FILD02", oRow, ""
        //			//        oDS_PH_PY111B.setValue "U_FILD03", oRow, 0
        //			//        oMat1.LoadFromDataSource
        //			//    End If

        //			oForm.Freeze(false);
        //			return;
        //			PH_PY111_AddMatrixRow_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY111_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region PH_PY111_Print_Report01
        //		private void PH_PY111_Print_Report01()
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
        //				goto PH_PY111_Print_Report01_Error;
        //			}

        //			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //			WinTitle = "[S142] 발주서";
        //			ReportName = "S142_1.rpt";
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			sQry = "EXEC PH_PY111_1 '" + oForm.Items.Item("8").Specific.Value + "'";
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
        //			PH_PY111_Print_Report01_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY111_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region PAY_Tax1(주석처리 되어 있음)
        ////Private Sub PAY_Tax1()
        ////On Error GoTo error_Message
        ////    Dim oRecordSet    As SAPbobsCOM.Recordset
        ////    Dim sQry          As String
        ////    Dim iCol          As Integer
        ////    Dim BnsUsed       As Boolean
        ////    Dim GABGUN As Double, JUMINN As Double, GBHAMT As Double, MEDAMT As Double, KUKAMT As Double
        ////    Dim INCOME As Double, GWASEE As Double, BTAX01 As Double, BTAX02 As Double, BTAX03 As Double, BTAX04 As Double, BTAX05 As Double, BTAX06 As Double, BTAX07 As Double
        ////    Dim MONPAY As Double, BT1AMT As Double, FODCSU As Double, CARCSU As Double, FRNCSU As Double, CHLCSU As Double, YGUCSU As Double, GITBTX As Double, GITBTX_N As Double
        ////    Dim BNSCSU As Double, BT1KUM As Double, TOTGBH As Double, TOTMED As Double
        ////    Dim KUKEND       As String, KUKSTR       As String
        ////    Dim U_NJYYMM As String, U_NJYRAT As Double, U_NJCRAT As Double, NJCMED As Double
        ////    Dim BASAMT As Double
        ////    Dim WK_YBTXAM(1 To 7) As Double
        ////    GABGUN = 0: JUMINN = 0: GBHAMT = 0: MEDAMT = 0: KUKAMT = 0: NJCMED = 0
        ////    INCOME = 0: GWASEE = 0: BTAX01 = 0: BTAX02 = 0: BTAX03 = 0: BTAX04 = 0: BTAX05 = 0: BTAX06 = 0: BTAX07 = 0:
        ////    MONPAY = 0: BT1AMT = 0: FODCSU = 0: CARCSU = 0: TOTGBH = 0: CHLCSU = 0: YGUCSU = 0: TOTMED = 0
        ////    BNSCSU = 0: GITBTX = 0: GITBTX_N = 0
        ////    BASAMT = 0  '/ 통상월급
        ////    U_NJYYMM = 0: U_NJYRAT = 0: U_NJCRAT = 0
        ////    BnsUsed = False
        ////    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        //// '/ 급여총액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        ////'/ 월정체크, 한도액, 생산비과대상자,부양가족공제,고용보험계산자
        ////    For iCol = 1 To 24
        ////        If WK_C.CSUCOD(iCol) = "A04" Then
        ////            If ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
        ////                WG03.U_CSUAMT(iCol) = WG03.U_BONUSS
        ////                BNSCSU = BNSCSU + Val(WG03.U_CSUAMT(iCol))
        ////                BnsUsed = True
        ////            Else
        ////                WG03.U_BONUSS = WG03.U_CSUAMT(iCol)
        ////            End If
        ////        End If
        ////        If Trim$(WK_C.GWATYP(iCol)) = "" Then WK_C.GWATYP(iCol) = "1"
        ////        Select Case WK_C.GWATYP(iCol)
        ////        Case "1"   '/ 과세
        ////            If WK_C.CSUKUM(iCol) > 0 And WG03.U_CSUAMT(iCol) > WK_C.CSUKUM(iCol) Then
        ////               WG03.U_CSUAMT(iCol) = Val(WK_C.CSUKUM(iCol))
        ////            End If
        ////        Case "2"   '/ 식대보조
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                FODCSU = 0
        ////            Else
        ////                FODCSU = Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////
        ////            If WK_C.CSUKUM(iCol) > 0 And FODCSU > WK_C.CSUKUM(iCol) Then
        ////                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX02 = BTAX02 + FODCSU
        ////            End If
        ////        Case "3"   '/ 차량보조
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                CARCSU = 0
        ////            Else
        ////                CARCSU = Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////
        ////            If WK_C.CSUKUM(iCol) > 0 And CARCSU > WK_C.CSUKUM(iCol) Then
        ////                BTAX02 = BTAX02 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX02 = BTAX02 + CARCSU
        ////            End If
        ////        Case "4"   '/ 생산비과
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                BT1AMT = BT1AMT + 0
        ////            Else
        ////                BT1AMT = BT1AMT + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////            If BT1KUM < WK_C.CSUKUM(iCol) Then BT1KUM = WK_C.CSUKUM(iCol)    '비과세4-마지막줄의 한도액만
        ////        Case "5"   '/ 국외비과세
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                FRNCSU = FRNCSU + 0
        ////            Else
        ////                FRNCSU = FRNCSU + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////
        ////            If WK_C.CSUKUM(iCol) > 0 And FRNCSU > WK_C.CSUKUM(iCol) Then
        ////                BTAX03 = BTAX03 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX03 = BTAX03 + FRNCSU
        ////            End If
        ////        Case "6"   '/ 연구수당
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                YGUCSU = YGUCSU + 0
        ////            Else
        ////                YGUCSU = YGUCSU + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////            If WK_C.CSUKUM(iCol) > 0 And YGUCSU > WK_C.CSUKUM(iCol) Then
        ////                BTAX05 = BTAX05 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX05 = BTAX05 + YGUCSU
        ////            End If
        ////        Case "7"   '/ 보육수당
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                CHLCSU = CHLCSU + 0
        ////            Else
        ////                CHLCSU = CHLCSU + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////            If WK_C.CSUKUM(iCol) > 0 And CHLCSU > WK_C.CSUKUM(iCol) Then
        ////                BTAX06 = BTAX06 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX06 = BTAX06 + CHLCSU
        ////            End If
        ////        Case "8"
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                GITBTX = 0
        ////            Else
        ////                GITBTX = Val(WG03.U_CSUAMT(iCol))  '/ 누적한도체크아님 각 수당별한도체크.
        ////            End If
        ////
        ////            If WK_C.CSUKUM(iCol) > 0 And GITBTX > WK_C.CSUKUM(iCol) Then
        ////                BTAX04 = BTAX04 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX04 = BTAX04 + GITBTX
        ////            End If
        ////        Case Else   '/ 비과세-기타-미제출
        ////            If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" Then
        ////                GITBTX_N = 0
        ////            Else
        ////                GITBTX_N = Val(WG03.U_CSUAMT(iCol))    '/ 누적한도체크아님 각 수당별한도체크.
        ////            End If
        ////            If WK_C.CSUKUM(iCol) > 0 And GITBTX_N > WK_C.CSUKUM(iCol) Then
        ////                BTAX07 = BTAX07 + WK_C.CSUKUM(iCol)
        ////            Else
        ////                BTAX07 = BTAX07 + GITBTX_N
        ////            End If
        ////        End Select
        ////        If oJOBTYP = "2" And WK_C.BNSUSE(iCol) = "Y" And ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
        ////            '/ 상여 상여금에 포함 수당은 총지급액에 합산하지않아야함.
        ////        Else
        ////            WG03.U_TOTPAY = WG03.U_TOTPAY + Val(WG03.U_CSUAMT(iCol))
        ////            '/ 월정급여
        ////            If WK_C.MPYGBN(iCol) = "Y" Then
        ////                MONPAY = MONPAY + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////            '/ 고용보험대상금액
        ////            If WK_C.GBHGBN(iCol) = "Y" Then
        ////                TOTGBH = TOTGBH + Val(WG03.U_CSUAMT(iCol))
        ////            End If
        ////        End If
        ////       '/ 통상임금
        ////        Select Case Trim$(WK_C.CSUCOD(iCol))
        ////        Case "A01", "E06", "E02", "E11", "E04", "C05" '/ (e06직책수당 + e02생산장려 + e11현장수당 + e04기능수당+C05주휴수당)/30일
        ////               BASAMT = BASAMT + WG03.U_CSUAMT(iCol)
        ////        End Select
        ////    Next iCol
        ////    If oJOBTYP = "2" Then
        ////        WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
        ////        MONPAY = MONPAY + WG03.U_BONUSS
        ////        TOTGBH = TOTGBH + WG03.U_BONUSS
        ////    End If
        ////    If oJOBTYP <> "1" And ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "Y" Then
        ////        Select Case oJOBTYP
        ////        Case "1"
        ////        Case "2"
        ////            WG03.U_BONUSS = WG03.U_TOTPAY
        ////            WG03.U_APPRAT = "100"
        ////            WG03.U_BNSRAT = "100"
        ////        Case "3"
        ////            WG03.U_APPRAT = "100"
        ////            WG03.U_BNSRAT = "100"
        ////        End Select
        ////    End If
        ////
        ////    '/ 상여총액(급여대장에 상여항목포함안될경우)
        ////'    If oJOBTYP = "2" And BnsUsed = False Then  '/ 상여만이면
        ////'        WG03.U_CSUAMT(1) = WG03.U_BONUSS
        ////'        BNSCSU = BNSCSU + Val(WG03.U_CSUAMT(1))
        ////'        WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
        ////'        '/ 고용보험대상금액
        ////'        If WK_C.GBHGBN(1) = "Y" Then
        ////'            TOTGBH = TOTGBH + WG03.U_BONUSS
        ////'        End If
        ////'        '/ 월정급여
        ////'        If WK_C.MPYGBN(1) = "Y" Then
        ////'            MONPAY = MONPAY + WG03.U_BONUSS
        ////'        End If
        ////'
        ////'    End If
        ////    WG03.U_TOTPAY = Format$(WG03.U_TOTPAY, "##########0")
        ////   If WG03.U_TOTPAY = 0 Then Exit Sub
        ////' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //// '/ 공제내역 계산/
        ////' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////'/ 세액계산 /*******************************************************************************/
        ////'/ 비과세산출 /*******************************************************************************/
        ////    '/ 비과세누적액 산출
        ////        WK_YBTXAM(1) = 0: WK_YBTXAM(2) = 0: WK_YBTXAM(3) = 0: WK_YBTXAM(4) = 0: WK_YBTXAM(5) = 0: WK_YBTXAM(6) = 0:: WK_YBTXAM(7) = 0
        ////        sQry = "EXEC PH_PY111_BTX '" & oYM & "','" & oJOBTYP & "', '" & oJOBGBN & "', '" & Trim$(WG03.U_MSTCOD) & "'"
        ////        oRecordSet.DoQuery sQry
        ////        If oRecordSet.RecordCount > 0 Then
        ////            WK_YBTXAM(1) = Val(oRecordSet.Fields(0).Value)   '/ 연누적액
        ////            WK_YBTXAM(2) = Val(oRecordSet.Fields(1).Value)   '/ 월누적
        ////            WK_YBTXAM(3) = Val(oRecordSet.Fields(2).Value)   '/ 월누적
        ////            WK_YBTXAM(4) = Val(oRecordSet.Fields(3).Value)   '/ 월누적
        ////            WK_YBTXAM(5) = Val(oRecordSet.Fields(4).Value)   '/ 월누적
        ////            WK_YBTXAM(6) = Val(oRecordSet.Fields(5).Value)   '/ 월누적
        ////            WK_YBTXAM(7) = Val(oRecordSet.Fields(6).Value)   '/ 월누적-기타비과세(미제출)
        ////        End If
        ////
        ////    '/ 1.1) 생산직 비과세
        ////    If sRecordset.Fields("U_BX1SEL").Value = "Y" Then             '/ 비과세대상자만
        ////       If (MONPAY - BT1AMT) <= 1000000 Then
        ////            BTAX01 = BT1AMT
        ////       Else
        ////            BTAX01 = 0
        ////       End If
        ////       If BT1KUM <> 0 And BTAX01 > BT1KUM Then BTAX01 = BT1KUM  '/ 한도액구함
        ////    '/ 연간 누적 비과세1 체크(상여계산시는 제외)
        ////       If oJOBTYP <> "2" Then
        ////          '/ (비과세1+연누적비과세)> 연240만원 THEN 비과세1 = 연240만원-(연누적비과세+비과세1)
        ////          If (BTAX01 + WK_YBTXAM(1)) > WG01.TB1KUM(1) Then
        ////             BTAX01 = WG01.TB1KUM(1) - WK_YBTXAM(1)
        ////          End If
        ////
        ////          If BTAX01 < 0 Then BTAX01 = 0
        ////       End If
        ////    End If
        ////
        ////    '/ 1.2) 과세대상급여 ( 총급여액-비과세소득 )
        ////    GWASEE = WG03.U_TOTPAY - BTAX01 - BTAX02 - BTAX03 - BTAX04 - BTAX05 - BTAX06
        ////    WG03.U_GWASEE = GWASEE
        ////    WG03.U_BTAX01 = BTAX01
        ////    WG03.U_BTAX02 = BTAX02
        ////    WG03.U_BTAX03 = BTAX03
        ////    WG03.U_BTAX04 = BTAX04
        ////    WG03.U_BTAX05 = BTAX05
        ////    WG03.U_BTAX06 = BTAX06
        ////    WG03.U_BTAX07 = BTAX07
        ////    '/ 1.3) 세액계산
        ////    '/ 세액계산안함으로 하면 소득세,주민세 계산안함.
        ////    If ods_PH_PY111A.getvalue("U_TAXCHK",0) =  "N" Then
        ////        '/ 총지급액, 과세대상금액, 공제인원수
        ////        GABGUN = 0
        ////        JUMINN = 0
        ////    Else
        ////        If sRecordset.Fields("U_TAXSEL").Value = "Y" Then
        ////            INCOME = GWASEE '- IIf(sRecordSet.Fields("U_BX1SEL").Value = "Y", BNSCSU, 0)
        ////            If oJOBTYP = "3" Then WG03.U_TAXTRM = WG03.U_TAXTRM + 1
        ////            If oJOBTYP <> "1" Then  '/ 상여처리
        ////               INCOME = MDC_SetMod.IInt((GWASEE + WG03.U_AVRPAY) / WG03.U_TAXTRM, 1)
        ////            End If
        ////         '/ 총지급액, 과세대상금액, 공제인원수
        ////            MDC_SetMod.Get_GabGunSe GABGUN, JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001
        ////        Else
        ////            GABGUN = 0
        ////            JUMINN = 0
        ////        End If
        ////    End If
        ////    '/ 1.4) 고용보험
        ////    '// 고용보험 계산안함으로 하면 제외
        ////    If ods_PH_PY111A.getvalue("U_GBHCHK",0) =  "N" Then
        ////            GBHAMT = 0
        ////    Else
        ////        If sRecordset.Fields("U_GBHSEL").Value = "Y" Then
        ////            GBHAMT = MDC_SetMod.IInt(TOTGBH * (0.45 / 100), 10)
        ////        End If
        ////    End If
        ////    '// 건강보험과 국민연금은 급여+정기, 급상여+정기일때만 자동계산
        ////    If (oJOBTYP = "1" And oJOBGBN = "1") Or (oJOBTYP = "3" And oJOBGBN = "1") Then
        ////        '/ 1.5) 의료보험 /'/ 당월 1일이후 입사자 건강보험료 제외
        ////        If Val(WG03.U_MEDAMT) <> 0 Then
        ////            sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO, U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] "
        ////            sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
        ////            oRecordSet.DoQuery sQry
        ////            If oRecordSet.RecordCount > 0 Then
        ////                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").Value
        ////                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").Value
        ////                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").Value
        ////                If Val(WG03.U_MEDAMT) < oRecordSet.Fields("U_FROM").Value Then
        ////                    MEDAMT = oRecordSet.Fields("U_FROM").Value
        ////                ElseIf oRecordSet.Fields("U_TO").Value > 0 And Val(WG03.U_MEDAMT) > oRecordSet.Fields("U_TO").Value Then
        ////                    MEDAMT = oRecordSet.Fields("U_TO").Value
        ////                Else
        ////                    MEDAMT = Val(WG03.U_MEDAMT)
        ////                End If
        ////                MEDAMT = MDC_SetMod.IInt(MEDAMT * Val(oRecordSet.Fields("U_EMPRAT").Value) / 100, 10)
        ////            End If
        ////            If oYM = Mid$(Replace(WG03.U_INPDAT, "-", ""), 1, 6) Then
        ////               If oYM & "01" <> Mid$(WG03.U_INPDAT, 1, 8) Then
        ////                    MEDAMT = 0                                 '/ 당월 1일이후 입사자 건강보험료 제외
        ////               End If
        ////            End If
        ////        End If
        ////        '/ 1.5.5) 노인장기요양보험료율
        ////        If MEDAMT <> 0 Then
        ////            sQry = "SELECT TOP 1 U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] "
        ////            sQry = sQry & " WHERE U_NJYYMM <= '" & oYM & "' ORDER BY U_NJYYMM DESC"
        ////            oRecordSet.DoQuery sQry
        ////            If oRecordSet.RecordCount > 0 Then
        ////                U_NJYYMM = oRecordSet.Fields("U_NJYYMM").Value
        ////                U_NJYRAT = oRecordSet.Fields("U_NJYRAT").Value
        ////                U_NJCRAT = oRecordSet.Fields("U_NJCRAT").Value
        ////            Else
        ////                U_NJYYMM = ""
        ////                U_NJYRAT = 0
        ////                U_NJCRAT = 0
        ////            End If
        ////        End If
        ////        '/ 1.6) 국민연금 /
        ////        If oYM <= "200803" Or Val(WG03.U_KUKAMT) = 0 Then
        ////            sQry = "SELECT TOP 1 ISNULL(T0.U_EMPAMT, 0) FROM [@ZPY102L] T0 WHERE T0.Code <= '" & oYM & "' "
        ////            sQry = sQry & " AND T0.U_CODNBR ='" & sRecordset.Fields("U_KUKGRD").Value & "'"
        ////            sQry = sQry & " ORDER BY CODE DESC"
        ////            oRecordSet.DoQuery sQry
        ////            If oRecordSet.RecordCount > 0 Then
        ////                KUKAMT = Val(oRecordSet.Fields(0).Value)
        ////            End If
        ////        Else
        ////            If Trim$(sRecordset.Fields("U_KUKGRD").Value) = "" And Val(WG03.U_KUKAMT) = 0 Then
        ////                KUKAMT = 0
        ////            Else
        ////                sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] "
        ////                sQry = sQry & " WHERE CODE <= '" & oYM & "' ORDER BY CODE DESC"
        ////                oRecordSet.DoQuery sQry
        ////                If oRecordSet.RecordCount > 0 Then
        ////                    If Val(WG03.U_KUKAMT) = 0 Then
        ////                        KUKAMT = 0
        ////                    ElseIf Val(WG03.U_KUKAMT) < oRecordSet.Fields("U_FROM").Value Then
        ////                        KUKAMT = oRecordSet.Fields("U_FROM").Value
        ////                    ElseIf oRecordSet.Fields("U_TO").Value > 0 And Val(WG03.U_KUKAMT) > oRecordSet.Fields("U_TO").Value Then
        ////                        KUKAMT = oRecordSet.Fields("U_TO").Value
        ////                    Else
        ////                        KUKAMT = Val(WG03.U_KUKAMT)
        ////                    End If
        ////                    KUKAMT = MDC_SetMod.IInt(KUKAMT * Val(oRecordSet.Fields("U_EMPRAT").Value) / 100, 10)
        ////                End If
        ////            End If
        ////        End If
        ////        '/ 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
        ////         If Trim$(WG03.U_PERNBR) <> "" Then
        ////             Select Case Mid$(Trim$(WG03.U_PERNBR), 7, 1)
        ////             Case "1", "2", "5", "6"
        ////                KUKSTR = Format$("19" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01", "0000-00-00")
        ////                KUKEND = Format$(DateAdd("yyyy", 60, KUKSTR), "yyyymm")
        ////             Case "3", "4", "7", "8"
        ////                KUKSTR = "20" & Mid$(Trim$(WG03.U_PERNBR), 1, 4) & "01"
        ////                KUKEND = Format$(DateAdd("y", 60, KUKSTR), "yyyymm")
        ////             Case Else
        ////                KUKSTR = ""
        ////                KUKEND = ""
        ////             End Select
        ////            If Trim$(KUKEND) <> "" And Trim$(KUKEND) < oYM Then
        ////                KUKAMT = 0
        ////            End If
        ////         End If
        ////'///////////////////////////////////////////////////
        ////    End If
        ////    If WG03.U_TAXTRM = 0 Then WG03.U_TAXTRM = 1
        ////'// 2. 공식가지고 계산하기-고정공제
        ////    If ods_PH_PY111A.getvalue("U_CHGCHK",0) =  "N" Then
        ////        For iCol = 1 To 18
        ////            Select Case WK_G.GONCOD(iCol)
        ////            Case "G01" '/ 소득세
        ////                If oJOBTYP = "1" Then
        ////                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt(GABGUN, 10)
        ////                Else
        ////                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10)
        ////                End If
        ////                If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
        ////            Case "G02" '/ 주민세
        ////                If oJOBTYP = "1" Then
        ////                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt(JUMINN, 10)
        ////                Else
        ////                    WG03.U_GONAMT(iCol) = MDC_SetMod.IInt((MDC_SetMod.IInt((GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, 10)) * 0.1, 10)
        ////                End If
        ////                If WG03.U_GONAMT(iCol) < 0 Then WG03.U_GONAMT(iCol) = 0
        ////            Case "G05" '/ 고용보험
        ////                WG03.U_GONAMT(iCol) = GBHAMT
        ////            Case "G03" '/ 국민연금
        ////                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + KUKAMT
        ////            Case "G04" '/ 건강보험
        ////                WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + MEDAMT
        ////                TOTMED = WG03.U_GONAMT(iCol)
        ////            Case "G09" '/ 노조회비
        ////                If oJOBTYP <> "2" Then    '/ 상여만 지급일경우 건강보험, 국민연금은 제외함.
        ////                    If Trim$(sRecordset.Fields("U_NOJGBN").Value) = "Y" Then
        ////                       ' WG03.U_GONAMT(iCol) = MDC_SetMod.RInt(WG03.U_TOTPAY * 0.015, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)) '2008.05월급여까지
        ////
        ////                    '/ '2008.07.15지급분(6월귀속)부터적용
        ////                         WG03.U_GONAMT(iCol) = MDC_SetMod.RInt(BASAMT * 0.02, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        ////                    Else
        ////                        WG03.U_GONAMT(iCol) = 0
        ////                    End If
        ////                End If
        ////            Case "G10" '/ 노인장기요양보험
        ////                If U_NJYYMM <= oYM Then
        ////                    NJCMED = MDC_SetMod.IInt(MEDAMT * (U_NJYRAT / 100), 10) '/ (10원미만 절사)
        ////                    If Trim$(WG03.U_NJCGBN) = "Y" Then '/ 경감대상자일경우
        ////                        NJCMED = NJCMED - (MDC_SetMod.IInt(NJCMED * (U_NJCRAT / 100) + 9, 10)) '/ 경감보험료(10원미만절상)
        ////                    End If
        ////                    WG03.U_GONAMT(iCol) = WG03.U_GONAMT(iCol) + NJCMED       '/ 장기요양보험료
        ////                End If
        ////            End Select
        ////        Next iCol
        ////    End If
        ////'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ////'/2. 총공제액
        ////'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ////    For iCol = 1 To 18
        ////        WG03.U_TOTGON = WG03.U_TOTGON + Val(WG03.U_GONAMT(iCol))
        ////    Next iCol
        ////'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ////'/3. 실지급액
        ////'/ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        ////    WG03.U_SILJIG = MDC_SetMod.IInt(WG03.U_TOTPAY - WG03.U_TOTGON, 1)
        ////   'If WG03.U_SILJIG < 1000 Then WG03.U_SILJIG = 0
        ////
        ////    Set oRecordSet = Nothing
        ////    Exit Sub
        ////'/////////////////////////////////////////////////////////////////////////////////////////////////
        ////error_Message:
        ////    Set oRecordSet = Nothing
        ////    Sbo_Application.StatusBar.SetText "PH_PY111_Save Error :" & Space$(10) & err.Description, bmt_Short, smt_Error
        ////End Sub
        #endregion

        #region PH_PY111_Validate
        //		public bool PH_PY111_Validate(string ValidateType)
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			functionReturnValue = true;
        //			object i = null;
        //			int j = 0;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY111A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY111A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				goto PH_PY111_Validate_Exit;
        //			}
        //			//
        //			if (ValidateType == "수정") {

        //			} else if (ValidateType == "행삭제") {

        //			} else if (ValidateType == "취소") {

        //			}
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY111_Validate_Exit:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY111_Validate_Error:
        //			functionReturnValue = false;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY111_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}
        #endregion

        #region Execution_Process (사용 안됨)
        //		private void Execution_Process()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string sQry = null;
        //			short ErrNum = 0;
        //			string oStrE = null;
        //			string MSTCOD = null;
        //			int oRow = 0;
        //			string Result = null;
        //			SAPbouiCOM.ComboBox oCombo = null;

        //			int oProValue = 0;
        //			int V_StatusCnt = 0;
        //			int TOTCNT = 0;
        //			//progbar
        //			//    oMat1.Clear
        //			MaxRow = 0;
        //			ErrNum = 0;
        //			REMARK1 = "";
        //			REMARK2 = "";
        //			REMARK3 = "";
        //			///1. 수당타이틀구함
        //			//UPGRADE_WARNING: Create_Tiltle 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oStrE = Create_Tiltle();
        //			/// 수당타이틀 생성
        //			if (!string.IsNullOrEmpty(oStrE)) {
        //				ErrNum = 1;
        //				goto Error_Message;
        //			}
        //			/// Question
        //			if (MDC_Globals.Sbo_Application.MessageBox("급여 계산을 실행하시겠습니까?", 2, "&Yes!", "&No") == 2) {
        //				ErrNum = 2;
        //				goto Error_Message;
        //			}
        //			sRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			/// 2.급여대상기간 구함.
        //			if (oJOBTRG == "%") {
        //				sQry = " SELECT T0.Code   FROM [@PH_PY107B] T0";
        //				sQry = sQry + " WHERE T0.CODE = (SELECT TOP 1 CODE FROM [@PH_PY107A] WHERE CODE <= '" + oYM + "' ORDER BY CODE DESC)";
        //				sQry = sQry + " GROUP BY T0.Code, T0.U_STRMON, T0.U_STRDAY, T0.U_JIGMON, T0.U_JIGDAY";
        //				sRecordset.DoQuery(sQry);
        //				//// 지급일자나 급여계산일자가 다를경우 전체로 할수없슴.
        //				if (sRecordset.RecordCount > 1) {
        //					ErrNum = 5;
        //					goto Error_Message;
        //				} else {
        //					oCombo = oForm.Items.Item("JOBTRG").Specific;
        //					////
        //					if (oCombo.ValidValues.Count > 1) {
        //						oJOBTRG = Strings.Trim(oCombo.ValidValues.Item(0).Value);
        //					}
        //					//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oCombo = null;
        //				}
        //			}

        //			/// 급여계산기간
        //			sQry = "SELECT DBO.Func_PAYTerm( '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oYM, "0000-00") + "-01" + "', '" + oJOBTRG + "') ";
        //			sRecordset.DoQuery(sQry);
        //			if (sRecordset.RecordCount == 0) {
        //				ErrNum = 3;
        //				goto Error_Message;
        //			} else {
        //				if (string.IsNullOrEmpty(Strings.Trim(sRecordset.Fields.Item(0).Value))) {
        //					ErrNum = 3;
        //					goto Error_Message;
        //				}
        //				StrDate = Strings.Mid(sRecordset.Fields.Item(0).Value, 1, 8);
        //				/// 급여계산시작일
        //				EndDate = Strings.Mid(sRecordset.Fields.Item(0).Value, 10, 8);
        //				/// 급여계산종료일
        //			}
        //			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oJOBTRG = Strings.Trim(oForm.Items.Item("JOBTRG").Specific.Selected.Value);

        //			/// 소득세 계산방법
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PAY_001 = MDC_SetMod.Get_ReData(ref "U_CODGBN", ref "CODE", ref "[@ZPY304H]", ref "'PAY001'", ref "");
        //			/// 상여금 평균임금 가져오는 방법
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PAY_007 = MDC_SetMod.Get_ReData(ref "U_CODGBN", ref "CODE", ref "[@ZPY304H]", ref "'PAY007'", ref "");
        //			/// 급여계산 처리 루틴
        //			if (Strings.Mid(oYM, 5, 2) == "12")
        //				JSNYER = Conversion.Val(Strings.Left(oYM, 4));
        //			else
        //				JSNYER = Conversion.Val(Strings.Left(oYM, 4)) - 1;
        //			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oMSTCOD = oForm.Items.Item("MSTCOD").Specific.String;
        //			/// 급여입퇴사자범위아닌사람 제거
        //			sQry = "DELETE FROM  [@PH_PY111A]";
        //			sQry = sQry + " FROM [@PH_PY111A] T0  INNER JOIN [OHEM] T1 ON T0.U_MSTCOD = T1.U_MSTCOD";
        //			sQry = sQry + " INNER JOIN [@PH_PY001A] T2 ON T0.U_MSTCOD = T2.Code";
        //			sQry = sQry + " WHERE   T0.U_YM = '" + oYM + "'";
        //			sQry = sQry + " AND     T0.U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND     T0.U_JOBGBN = '" + oJOBGBN + "'";
        //			sQry = sQry + " AND     (T2.U_JOBTRG LIKE '" + oJOBTRG + "' OR ISNULL(T2.U_JOBTRG, 0) = '0')";
        //			sQry = sQry + " AND     T0.U_ENDCHK <> 'Y'";
        //			//잠금아닌것만
        //			sQry = sQry + " AND     (CONVERT(CHAR(8), T1.StartDate,112) > '" + EndDate + "'";
        //			sQry = sQry + "            OR (ISNULL(CONVERT(CHAR(8),T1.TermDate,112),'') <> ''";
        //			sQry = sQry + "                 AND CONVERT(CHAR(8), T1.TermDate, 112) < '" + StrDate + "'))";
        //			/// 당월퇴사자 제외
        //			if (Strings.Trim(oDS_PH_PY111A.GetValue("U_JOBTRG", 0)) != "%") {
        //				sQry = sQry + "   AND ISNULL(T2.U_JOBTRG, 0) = '" + oJOBTRG + "'";
        //				//급여지급일
        //			}
        //			if (Strings.Trim(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) != "%") {
        //				sQry = sQry + "   AND ISNULL(T0.U_CLTCOD, 0) = '" + oCLTCOD + "'";
        //				//자사코드
        //			}
        //			sQry = sQry + " AND     T0.U_MSTDPT BETWEEN '" + oSTRDPT + "' AND '" + oENDDPT + "'";
        //			if (oMSTCOD != "%" & !string.IsNullOrEmpty(oMSTCOD)) {
        //				sQry = sQry + " AND     T0.U_MSTCOD LIKE '" + oMSTCOD + "'";
        //				//사원번호
        //			}
        //			sRecordset.DoQuery(sQry);

        //			///3. 인사마스터 조회-대상자 조회
        //			sQry = " SELECT  T0.U_MSTCOD, (T0.LastName + T0.FirstName) AS U_MSTNAM, T0.EmpID, T0.U_CLTCOD, T0.Branch AS U_MSTBRK, ";
        //			sQry = sQry + " T3.U_MSTDPT, T4.U_MSTSTP,  ISNULL(T6.U_CodeNm,'') AS U_CLTNAM, T2.Name AS U_BRKNAM, T3.Name AS U_DPTNAM, T4.Name AS U_STPNAM, T1.U_PAYTYP,";
        //			sQry = sQry + " T1.U_JIGCOD, T1.U_HOBONG, T1.U_STDAMT, T1.U_BNSAMT, CONVERT(CHAR(8), T0.StartDate,112) AS U_INPDAT, ";
        //			sQry = sQry + "  CONVERT(CHAR(8), T0.TermDate, 112) AS U_OUTDAT, ISNULL(T1.U_GBHSEL, 'N') AS U_GBHSEL, T1.U_KUKGRD, T1.U_MEDAMT, ISNULL(T1.U_KUKAMT,0) AS U_KUKAMT, ISNULL(T1.U_GBHAMT,0) AS U_GBHAMT, ";
        //			sQry = sQry + " (1+ ISNULL(T1.U_BAEWOO,0)+ISNULL(T1.U_BUYNSU,0)) AS U_TAXCNT, ISNULL(T1.U_DAGYSU,0) AS U_CHLCNT,T1.U_TAXSEL, ISNULL(T1.U_FRGSEL, 'N') AS U_FRGSEL, ISNULL(T1.U_KUKOVR,'N') AS U_KUKOVR, ";
        //			sQry = sQry + " ISNULL(T1.U_BX1SEL, 'N') AS U_BX1SEL , ISNULL(T1.U_BNSSEL, 'N') AS U_BNSSEL, T0.U_INPRAT, T0.U_GNTEXE, T0.U_INPGBN, T0.U_NOJGBN, ISNULL(T1.U_AVRAMT,0) AS U_AVRAMT,";
        //			sQry = sQry + " CONVERT(CHAR(8), T0.U_INEDAT, 112) AS U_INEDAT, CONVERT(CHAR(8), T0.U_GRPDAT, 112) AS U_GRPDAT, ISNULL(T1.U_NJCGBN,'N') AS U_NJCGBN,ISNULL(T1.U_MEDFRG, '0') AS U_MEDFRG, ";
        //			sQry = sQry + " ISNULL(T5.U_CHAGAB, 0) AS U_CHAGAB, ISNULL(T5.U_CHANON,0) AS U_CHANON, ISNULL(T5.U_CHAJUM,0) AS U_CHAJUM, ISNULL(REPLACE(T0.GovID,'-',''),'') AS U_PERNBR, T1.U_JOBTRG ";
        //			sQry = sQry + " FROM [OHEM] T0 INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code";
        //			sQry = sQry + "                    LEFT JOIN [OUBR] T2 ON T0.Branch = T2.Code";
        //			sQry = sQry + "                    LEFT JOIN [OUDP] T3 ON T0.Dept = T3.Code";
        //			sQry = sQry + "                    LEFT JOIN [OHPS] T4 ON T0.position = T4.posID";
        //			sQry = sQry + "                    LEFT JOIN [@ZPY504H] T5 ON T0.U_MSTCOD = T5.U_MSTCOD AND ISNULL(T5.U_JSNYER, '') = '" + Strings.Trim(Convert.ToString(JSNYER)) + "' AND ISNULL(T5.U_JSNGBN,'') = '1' AND T0.U_CLTCOD = T5.U_CLTCOD";
        //			/// 연말정산
        //			sQry = sQry + "                    LEFT JOIN [@PS_HR200L] T6 ON T0.U_CLTCOD = T6.U_CODE AND T6.CODE = 'P144'";
        //			sQry = sQry + " WHERE CONVERT(CHAR(8), T0.StartDate,112) <= '" + EndDate + "'";
        //			sQry = sQry + "   AND (ISNULL(CONVERT(CHAR(8),T0.TermDate,112),'') ='" + Strings.Space(0) + "'";
        //			sQry = sQry + "    OR CONVERT(CHAR(8), T0.TermDate, 112) >= '" + StrDate + "')";
        //			/// 당월퇴사자 제외
        //			sQry = sQry + "   AND ISNULL(T1.U_JOBTRG, 0) <> '0'";
        //			/// 급여지급제외자"
        //			if (oJOBTRG != "%") {
        //				sQry = sQry + "   AND ISNULL(T1.U_JOBTRG, 0) = '" + oJOBTRG + "'";
        //				//급여지급일
        //			}
        //			if (Strings.Trim(oDS_PH_PY111A.GetValue("U_CLTCOD", 0)) != "%") {
        //				sQry = sQry + "   AND ISNULL(T0.U_CLTCOD, 0) = '" + oCLTCOD + "'";
        //				//자사코드
        //			}
        //			sQry = sQry + "   AND T3.U_MSTDPT BETWEEN '" + oSTRDPT + "' AND '" + oENDDPT + "'";
        //			if (oMSTCOD != "%" & !string.IsNullOrEmpty(oMSTCOD)) {
        //				sQry = sQry + "   AND ISNULL(T0.U_MSTCOD, '') = '" + oMSTCOD + "'";
        //				//사원번호
        //			}
        //			sQry = sQry + "   ";
        //			//정산년도
        //			sQry = sQry + " ORDER BY T0.U_CLTCOD, T0.Branch, T3.U_MSTDPT, T4.U_MSTSTP, T0.U_MSTCOD";
        //			sRecordset.DoQuery(sQry);
        //			if (sRecordset.RecordCount == 0) {
        //				ErrNum = 4;
        //				goto Error_Message;
        //			}
        //			////
        //			if ((MDC_Globals.oProgBar != null)) {
        //				MDC_Globals.oProgBar.Stop();
        //				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				MDC_Globals.oProgBar = null;
        //			}

        //			MDC_Globals.oProgBar = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 100, false);
        //			//// Process /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			sRecordset.MoveLast();
        //			G_TotCnt = sRecordset.RecordCount;
        //			/// 총급여처리대상인원수
        //			G_PayCnt = 0;
        //			G_ChkCnt = 0;
        //			oRow = 0;
        //			/// ProgressBar 메모리 할당
        //			TOTCNT = G_TotCnt;
        //			V_StatusCnt = System.Math.Round(TOTCNT / 100, 0);
        //			oProValue = 1;
        //			sRecordset.MoveFirst();
        //			while (!(sRecordset.EoF)) {
        //				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				MSTCOD = sRecordset.Fields.Item("U_MSTCOD").Value;
        //				oRow = oRow + 1;
        //				/// 초기화
        //				PAY_Initial();
        //				/// 급여계산
        //				Basis_Setting();

        //				if (oJOBTYP != "2" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "N") {
        //					//            Select Case Trim$(MDC_COMpanyGubun)
        //					//            'Case "CL", "HS", "SO"
        //					//            Case "HS"
        //					//                Call PayRoll_Process1
        //					//            Case Else    '/ 로직사용
        //					Result = PayRoll_Process();
        //					if (!string.IsNullOrEmpty(Strings.Trim(Result))) {
        //						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD + ": " + sRecordset.Fields.Item("U_MSTNAM").Value + Strings.Space(1) + Result;
        //						//                    oMat1.AddRow
        //						//                    ErrNum = 9
        //						//                    GoTo error_Message
        //					}
        //					//            End Select
        //				}
        //				/// 상여계산
        //				/// 상여지급여부
        //				if (oJOBTYP != "1" & sRecordset.Fields.Item("U_BNSSEL").Value == "Y" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "N") {
        //					Result = BonusRoll_Process();
        //					if (!string.IsNullOrEmpty(Strings.Trim(Result))) {
        //						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD + ": " + sRecordset.Fields.Item("U_MSTNAM").Value + Strings.Space(1) + Result;
        //						//                    oMat1.AddRow
        //						//                ErrNum = 8
        //						//                GoTo error_Message
        //					}
        //				}
        //				/// 변동수당/공제항목 추가
        //				ChangeItem_Process();
        //				/// 세액계산
        //				//       Select Case Trim$(MDC_COMpanyGubun)
        //				//       Case "CL", "HS", "SO"
        //				//            Call PAY_Tax1
        //				//       Case Else
        //				PAY_Tax();
        //				//       End Select
        //				/// 연말정산징수환급해주세용
        //				if (Strings.Trim(oDS_PH_PY111A.GetValue("U_JSNCHK", 0)) == "Y") {
        //					JeongSan_Process();
        //				}
        //				/// 급상여자료 저장
        //				if (WG03.U_TOTPAY != 0 | WG03.U_TOTGON != 0) {
        //					oForm.DataSources.UserDataSources.Item("Col0").Value = Convert.ToString(oRow);
        //					Result = PH_PY111_Save();
        //					if (Strings.Trim(Result) == "Y") {
        //						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD + ": " + sRecordset.Fields.Item("U_MSTNAM").Value + " 급여계산 완료.";
        //					} else {
        //						oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD + ": " + sRecordset.Fields.Item("U_MSTNAM").Value + Strings.Space(1) + Result;
        //					}
        //				} else {
        //					oForm.DataSources.UserDataSources.Item("Col1").Value = MSTCOD + ": " + sRecordset.Fields.Item("U_MSTNAM").Value + " 급여총지급액이 0입니다.";
        //				}
        //				//        oMat1.AddRow

        //				//// 상태보여주기
        //				if ((TOTCNT > 100 & oRow == oProValue * V_StatusCnt) | TOTCNT <= 100) {
        //					MDC_Globals.oProgBar.Text = oRow + "/ " + TOTCNT + " 건 처리중...!";
        //					oProValue = oProValue + 1;
        //					MDC_Globals.oProgBar.Value = oProValue;
        //				}


        //				sRecordset.MoveNext();
        //			}
        //			/// END
        //			MDC_Globals.oProgBar.Stop();
        //			sQry = "총: " + G_TotCnt + "( 처리 :" + G_PayCnt + " 잠김 :" + G_ChkCnt + ")건이 처리되었습니다.";
        //			if (!string.IsNullOrEmpty(Strings.Trim(REMARK1)))
        //				sQry = sQry + Constants.vbCrLf + "고용보험 나이제외:" + REMARK1;
        //			if (!string.IsNullOrEmpty(Strings.Trim(REMARK2)))
        //				sQry = sQry + Constants.vbCrLf + "국민연금 나이제외:" + REMARK2;
        //			if (!string.IsNullOrEmpty(Strings.Trim(REMARK3)))
        //				sQry = sQry + Constants.vbCrLf + "실지급액이 마이너스 사원:" + REMARK3;
        //			if (TermCHK == true)
        //				sQry = sQry + Constants.vbCrLf + "상여세액기간에 급여자료가 없습니다.세액이 적게 나올 수 있습니다.";

        //			MDC_Globals.Sbo_Application.MessageBox(sQry);
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("총:" + G_TotCnt + "( 처리 :" + G_PayCnt + " 잠김 :" + G_ChkCnt + ")건이 처리되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //			/// ProgressBar 중지/메모리 반환
        //			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			sRecordset = null;
        //			//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			MDC_Globals.oProgBar = null;
        //			return;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			sRecordset = null;
        //			if ((MDC_Globals.oProgBar != null)) {
        //				MDC_Globals.oProgBar.Stop();
        //				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				MDC_Globals.oProgBar = null;
        //			}

        //			if (ErrNum == 1) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText(oStrE, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			} else if (ErrNum == 2) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("급여계산작업이 취소되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //			} else if (ErrNum == 3) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("(ZPY110)급여대상기준일을 먼저 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			} else if (ErrNum == 4) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("지급 대상자가 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			} else if (ErrNum == 5) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("지급일 또는 급여기간이 다를 경우 지급 대상자 구분을 모두로 선택할수 없습니다. 확인해주세요", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //			} else if (ErrNum == 8) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("Bonus Error: 사번:" + MSTCOD + Result, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			} else if (ErrNum == 9) {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("PayRoll Error: 사번:" + MSTCOD + Result, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			} else {
        //				MDC_Globals.Sbo_Application.StatusBar.SetText("Execution_Process Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //			}
        //		}
        #endregion

        #region PAY_Initial (Execution_Process 에서 실행됨)
        //		private void PAY_Initial()
        //		{
        //			short iCol = 0;
        //			string GNSGBN = null;

        //			GNSGBN = "";
        //			WG03.DocNum = 0;
        //			//문서번호
        //			WG03.U_BASAMT = 0;
        //			WG03.U_DAYAMT = 0;
        //			WG03.U_MSTCOD = Strings.Trim(sRecordset.Fields.Item("U_MSTCOD").Value);
        //			WG03.U_MSTNAM = Strings.Trim(sRecordset.Fields.Item("U_MSTNAM").Value);
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_EmpID = sRecordset.Fields.Item("EmpID").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_CLTCOD = sRecordset.Fields.Item("U_CLTCOD").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_MSTBRK = sRecordset.Fields.Item("U_MSTBRK").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_MSTDPT = sRecordset.Fields.Item("U_MSTDPT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_MSTSTP = sRecordset.Fields.Item("U_MSTSTP").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_CLTNAM = sRecordset.Fields.Item("U_CLTNAM").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_BRKNAM = sRecordset.Fields.Item("U_BRKNAM").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_DPTNAM = sRecordset.Fields.Item("U_DPTNAM").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_STPNAM = sRecordset.Fields.Item("U_STPNAM").Value;
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_PAYTYP = MDC_SetMod.Get_ReData(ref "U_RelCd", ref "U_Minor", ref "[@PS_HR200L]", ref "'" + Strings.Trim(sRecordset.Fields.Item("U_PAYTYP").Value) + "'", ref " AND Code='P132'");
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_JOBTRG = sRecordset.Fields.Item("U_JOBTRG").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_PERNBR = sRecordset.Fields.Item("U_PERNBR").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_JIGCOD = sRecordset.Fields.Item("U_JIGCOD").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_HOBONG = sRecordset.Fields.Item("U_HOBONG").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_STDAMT = sRecordset.Fields.Item("U_STDAMT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_MEDAMT = sRecordset.Fields.Item("U_MEDAMT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_KUKAMT = sRecordset.Fields.Item("U_KUKAMT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_GBHAMT = sRecordset.Fields.Item("U_GBHAMT").Value;

        //			/// 입사일자(근속일기준: 1-그룹입사일기준, 2-입사일기준)
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			GNSGBN = MDC_SetMod.Get_ReData(ref "U_GNSGBN", ref "U_YM", ref "[@PH_PY106A]", ref "'" + Strings.Trim(U_CSUCOD) + "'", ref " AND U_PAYTYP='" + Strings.Trim(WG03.U_PAYTYP) + "'");
        //			/// 그룹입사일기준
        //			if (Strings.Trim(GNSGBN) == "1") {
        //				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG03.U_INPDAT = sRecordset.Fields.Item("U_GRPDAT").Value;
        //			} else {
        //				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG03.U_INPDAT = sRecordset.Fields.Item("U_INPDAT").Value;
        //			}
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_INEDAT = sRecordset.Fields.Item("U_INEDAT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_OUTDAT = sRecordset.Fields.Item("U_OUTDAT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_TAXCNT = sRecordset.Fields.Item("U_TAXCNT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_CHLCNT = sRecordset.Fields.Item("U_CHLCNT").Value;
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_NJCGBN = sRecordset.Fields.Item("U_NJCGBN").Value;
        //			WG03.U_MEDFRG = Conversion.Val(sRecordset.Fields.Item("U_MEDFRG").Value);
        //			//// 수당항목
        //			for (iCol = 1; iCol <= 24; iCol++) {
        //				WG03.U_CSUAMT[iCol] = 0;
        //			}
        //			WG03.U_GWASEE = 0;
        //			WG03.U_BTAX01 = 0;
        //			WG03.U_BTAX02 = 0;
        //			WG03.U_BTAX03 = 0;
        //			WG03.U_BTAX04 = 0;
        //			WG03.U_BTAX05 = 0;
        //			WG03.U_BTAX06 = 0;
        //			WG03.U_BTAX07 = 0;

        //			WG03.U_BTXG01 = 0;
        //			WG03.U_BTXH01 = 0;
        //			WG03.U_BTXH05 = 0;
        //			WG03.U_BTXH06 = 0;
        //			WG03.U_BTXH07 = 0;
        //			WG03.U_BTXH08 = 0;
        //			WG03.U_BTXH09 = 0;
        //			WG03.U_BTXH10 = 0;
        //			WG03.U_BTXH11 = 0;
        //			WG03.U_BTXH12 = 0;
        //			WG03.U_BTXH13 = 0;
        //			WG03.U_BTXI01 = 0;
        //			WG03.U_BTXK01 = 0;
        //			WG03.U_BTXM01 = 0;
        //			WG03.U_BTXM02 = 0;
        //			WG03.U_BTXM03 = 0;
        //			WG03.U_BTXO01 = 0;
        //			WG03.U_BTXQ01 = 0;
        //			WG03.U_BTXR10 = 0;
        //			WG03.U_BTXS01 = 0;
        //			WG03.U_BTXT01 = 0;
        //			WG03.U_BTXX01 = 0;
        //			WG03.U_BTXY01 = 0;
        //			WG03.U_BTXY02 = 0;
        //			WG03.U_BTXY03 = 0;
        //			WG03.U_BTXY20 = 0;
        //			WG03.U_BTXY21 = 0;
        //			WG03.U_BTXY22 = 0;
        //			WG03.U_BTXZ01 = 0;
        //			WG03.U_BTXTOT = 0;

        //			WG03.U_TOTPAY = 0;
        //			//// 공제항목
        //			for (iCol = 1; iCol <= 18; iCol++) {
        //				WG03.U_GONAMT[iCol] = 0;
        //			}

        //			WG03.U_TOTGON = 0;
        //			WG03.U_SILJIG = 0;
        //			//// 상여금
        //			WG03.U_AVRPAY = 0;
        //			WG03.U_NABTAX = 0;
        //			WG03.U_BNSRAT = 0;
        //			WG03.U_APPRAT = 0;
        //			WG03.U_GNSYER = 0;
        //			WG03.U_GNSMON = 0;
        //			WG03.U_TAXTRM = 0;
        //			WG03.U_BONUSS = 0;
        //			//// 2011.10.05 상여금 평균임금 급여기본등록에서 가져오기
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WK_AVRAMT = sRecordset.Fields.Item("U_AVRAMT").Value;
        //			//// 2009년추가 비과세코드가져오기
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TB1_BT3COD = MDC_SetMod.Get_ReData(ref "MAX(U_BTXCOD)", ref "U_GWATYP", ref "[@ZPY107L]", ref "'5'", ref " AND Left(Code,4) <='" + Strings.Left(oYM, 4) + "' GROUP BY CODE ORDER BY CODE DESC");
        //			if (string.IsNullOrEmpty(Strings.Trim(TB1_BT3COD)))
        //				TB1_BT3COD = "M01";

        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TB1_BT5COD = MDC_SetMod.Get_ReData(ref "MAX(U_BTXCOD)", ref "U_GWATYP", ref "[@ZPY107L]", ref "'6'", ref " AND Left(Code,4) <='" + Strings.Left(oYM, 4) + "' GROUP BY CODE ORDER BY CODE DESC");
        //			if (string.IsNullOrEmpty(Strings.Trim(TB1_BT5COD)))
        //				TB1_BT5COD = "H10";

        //		}
        #endregion

        #region PayRoll_Process (Execution_Process 에서 실행됨)
        //		private string PayRoll_Process()
        //		{
        //			string functionReturnValue = null;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			SAPbobsCOM.Recordset pRecordset = null;
        //			string sQry = null;
        //			short ErrNum = 0;

        //			short iCol = 0;
        //			int kCol = 0;
        //			string SuSilStr = null;
        //			double Tmp_CSUAMT = 0;

        //			ErrNum = 0;
        //			//// 각종수당들은 해당이름으로 변수만들어서 저장, 계산식 만드는것은 나름 저장그래야 루프문안씀
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			pRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//// 1.1. 공식가져오기-수당
        //			sQry = " SELECT T0.U_LINSEQ, T0.U_CSUCOD, T0.U_SILCUN, T0.U_SILCOD";
        //			sQry = sQry + " FROM [@PH_PY106B] T0 INNER JOIN [@PH_PY106A] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE   T1.U_YM = '" + Strings.Trim(U_CSUCOD) + "'";
        //			sQry = sQry + " AND     T1.U_PAYTYP = '" + Strings.Trim(WG03.U_PAYTYP) + "'";
        //			sQry = sQry + " ORDER BY T0.U_LINSEQ, T0.U_CSUCOD";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				ErrNum = 1;
        //				goto Error_Message;
        //			}
        //			while (!(oRecordSet.EoF)) {
        //				iCol = Conversion.Val(oRecordSet.Fields.Item("U_LINSEQ").Value);
        //				if (iCol > 0 & !string.IsNullOrEmpty(Strings.Trim(oRecordSet.Fields.Item("U_CSUCOD").Value))) {
        //					//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WK_C.GONSIL[iCol] = oRecordSet.Fields.Item("U_SILCUN").Value;
        //					/// 급여계산식
        //				} else {
        //					if (Strings.Left(oRecordSet.Fields.Item("U_CSUCOD").Value, 1) == "X") {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SuSilStr = oRecordSet.Fields.Item("U_SILCUN").Value;
        //						/// 급여계산식
        //						/// 2.2. 공식안 시스템제공코드값으로 변경
        //						SuSilStr = Change_GOSIL(ref SuSilStr);
        //						switch (Strings.Trim(oRecordSet.Fields.Item("U_CSUCOD").Value)) {
        //							case "X01":
        //								/// 기본일급
        //								X01_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X02":
        //								/// 통상일급
        //								X02_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X03":
        //								/// 기본시급
        //								X03_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X04":
        //								/// 통상시급
        //								X04_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //						}
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}
        //			WG03.U_DAYAMT = X01_Val;
        //			WG03.U_BASAMT = X02_Val;

        //			//// 2. 공식가지고 계산하기-수당
        //			for (iCol = 1; iCol <= 24; iCol++) {
        //				/// 수당항목이 있는것만
        //				if (!string.IsNullOrEmpty(Strings.Trim(WK_C.CSUCOD[iCol]))) {
        //					/// 공식이있으면 공식계산
        //					if (!string.IsNullOrEmpty(Strings.Trim(WK_C.GONSIL[iCol]))) {
        //						/// 상여만아니면
        //						if (oJOBTYP == "1" | oJOBTYP == "3") {
        //							/// 2.1. 공식-계산결과값가져오는거면..
        //							SuSilStr = WK_C.GONSIL[iCol];
        //							/// 계산된값일 경우
        //							for (kCol = 1; kCol <= 24; kCol++) {
        //								SuSilStr = Strings.Replace(SuSilStr, "#" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(kCol, "00"), Convert.ToString(WG03.U_CSUAMT[kCol]));
        //							}
        //							/// 2.2. 공식안 시스템제공코드값으로 변경
        //							SuSilStr = Change_GOSIL(ref SuSilStr);
        //							/// 2.3. 공식계산하기
        //							Tmp_CSUAMT = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //							/// 2.4. 정답가져오기
        //							//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref Tmp_CSUAMT, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //							if (Strings.Trim(WK_C.CSUCOD[iCol]) == "A04" & (oJOBTYP == "1" | sRecordset.Fields.Item("U_BNSSEL").Value != "Y")) {
        //								WG03.U_CSUAMT[iCol] = 0;
        //							}
        //						}
        //					}
        //				}
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			pRecordset = null;
        //			functionReturnValue = "";
        //			return functionReturnValue;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			pRecordset = null;
        //			if (ErrNum == 1) {
        //				functionReturnValue = "수당계산식이 없습니다. 확인하여 주십시오. (적용연월:" + U_CSUCOD + "급여형태: " + WG03.U_PAYTYP + ")";
        //			} else {
        //				functionReturnValue = Err().Number + Strings.Space(10) + Err().Description;
        //			}
        //			return functionReturnValue;
        //		}
        #endregion

        #region Basis_Setting (Execution_Process 에서 실행됨)
        //		private void Basis_Setting()
        //		{
        //			string STRDAT = null;
        //			string ENDDAT = null;
        //			///
        //			X01_Val = 0;
        //			X02_Val = 0;
        //			X03_Val = 0;
        //			X04_Val = 0;
        //			X10_Val = 0;
        //			X12_Val = 0;
        //			X13_Val = 0;
        //			X14_Val = "";
        //			X15_Val = 0;
        //			X16_Val = 0;
        //			X17_Val = 0;
        //			X18_Val = 0;
        //			X19_Val = 0;
        //			X20_Val = 0;
        //			//// X01,X02,X03,X04 <---수식초기에 셋팅


        //			//// X10:종료일수(해당작업기간의 월총일수 가져오기)
        //			//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //			X10_Val = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StrDate, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(EndDate, "0000-00-00"))) + 1;
        //			/// 한샘, 케어라인


        //			//// X11, X12, X13: 근속년수/근속월수/근속일수 가져오기
        //			if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_OUTDAT)) & WG03.U_OUTDAT <= EndDate) {
        //				MDC_SetMod.Term2(ref WG03.U_INPDAT, ref WG03.U_OUTDAT);
        //			} else {
        //				MDC_SetMod.Term2(ref WG03.U_INPDAT, ref EndDate);
        //			}
        //			X11_Val = MDC_Globals.ZPAY_GBL_GNSYER;
        //			X12_Val = MDC_Globals.ZPAY_GBL_GNSMON;
        //			X13_Val = MDC_Globals.ZPAY_GBL_GNSDAY;

        //			//// X14: 당월입퇴사유무
        //			if ((WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate) | (WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate)) {
        //				X14_Val = "1";
        //			} else {
        //				X14_Val = "0";
        //			}

        //			//// X15:총지급액 WG03.TOTPAY

        //			//// X16: 입퇴사일수
        //			X16_Val = X10_Val;
        //			if ((WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate) | (WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate)) {
        //				switch (true) {
        //					case (WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate) & (WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate):
        //						/// 당월입퇴사자
        //						MDC_SetMod.Term2(ref WG03.U_INPDAT, ref WG03.U_OUTDAT);
        //						break;
        //					case WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate:
        //						/// 당월입사자
        //						MDC_SetMod.Term2(ref WG03.U_INPDAT, ref EndDate);
        //						break;
        //					case WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate:
        //						/// 당월퇴사자
        //						MDC_SetMod.Term2(ref StrDate, ref WG03.U_OUTDAT);
        //						break;
        //				}
        //				X16_Val = (MDC_Globals.ZPAY_GBL_GNSMON * X10_Val + MDC_Globals.ZPAY_GBL_GNSDAY);
        //			}

        //			//// X18:수습일 (당월수습기간에 해당하는지 여부)
        //			X18_Val = 0;
        //			if ((WG03.U_INEDAT >= StrDate | (WG03.U_INEDAT <= EndDate & StrDate < WG03.U_INEDAT))) {
        //				if (StrDate <= WG03.U_INPDAT)
        //					STRDAT = WG03.U_INPDAT;
        //				else
        //					STRDAT = StrDate;
        //				if (EndDate <= WG03.U_INEDAT)
        //					ENDDAT = EndDate;
        //				else
        //					ENDDAT = WG03.U_INEDAT;
        //				if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_OUTDAT))) {
        //					if (WG03.U_OUTDAT <= EndDate & WG03.U_OUTDAT <= WG03.U_INEDAT) {
        //						ENDDAT = WG03.U_OUTDAT;
        //					}
        //				}
        //				MDC_SetMod.Term2(ref STRDAT, ref ENDDAT);
        //				X18_Val = (MDC_Globals.ZPAY_GBL_GNSMON * X10_Val + MDC_Globals.ZPAY_GBL_GNSDAY);
        //			}

        //			//// X17:수습율
        //			X17_Val = Conversion.Val(sRecordset.Fields.Item("U_INPRAT").Value) / 100;
        //			/// 수습율
        //			if (Conversion.Val(Convert.ToString(X17_Val)) == 0)
        //				X17_Val = 1;
        //			///수습율입력안하면 기본급 100%

        //			//// X19:근속일수(급여종료일기준)
        //			if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_OUTDAT)) & WG03.U_OUTDAT <= EndDate) {
        //				//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //				X19_Val = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_OUTDAT, "0000-00-00"))) + 1;
        //			} else {
        //				//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //				X19_Val = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(EndDate, "0000-00-00"))) + 1;
        //			}

        //			//// X20:근속일수(상여종료일기준) --2009.10.22 수정함 퇴사일반영.
        //			if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_OUTDAT)) & WG03.U_OUTDAT <= oGNEDAT) {
        //				//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //				X20_Val = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_OUTDAT, "0000-00-00"))) + 1;
        //			} else {
        //				//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //				X20_Val = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oGNEDAT, "0000-00-00"))) + 1;
        //			}
        //		}
        #endregion

        #region PayRoll_Process1
        //		private void PayRoll_Process1()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			SAPbobsCOM.Recordset pRecordset = null;
        //			string sQry = null;
        //			short iCol = 0;
        //			double INPRAT = 0;
        //			double ENDDAY = 0;
        //			string INPCheck = null;
        //			double Imsi_DAYAMT = 0;
        //			double Imsi_TIMAMT = 0;
        //			//// 각종수당들은 해당이름으로 변수만들어서 저장, 계산식 만드는것은 나름 저장그래야 루프문안씀
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			pRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			///2. 기본급여사항가져오기.
        //			/// 고정수당
        //			sQry = " SELECT U_FILD01, U_FILD02, U_FILD03 FROM [@ZPY127L] WHERE CODE  ='" + Strings.Trim(WG03.U_MSTCOD) + "' ";
        //			sQry = sQry + "  AND ISNULL(U_FILD01, '') <> '' ";
        //			//AND ISNULL(U_FILD03, 0) <> 0"
        //			sQry = sQry + " ORDER BY CODE, LineID";
        //			oRecordSet.DoQuery(sQry);
        //			while (!(oRecordSet.EoF)) {
        //				for (iCol = 1; iCol <= 24; iCol++) {
        //					if (oRecordSet.Fields.Item("U_FILD01").Value == WK_C.CSUCOD[iCol]) {
        //						//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref (oRecordSet.Fields.Item("U_FILD03").Value), WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //					} else if (WK_C.CSUCOD[iCol] == "A01") {
        //						//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref (sRecordset.Fields.Item("U_STDAMT").Value), WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}
        //			//    '/ 고정공제
        //			//    sQry = " SELECT U_FILD01, U_FILD02, U_FILD03 FROM [@ZPY127M] WHERE CODE  ='" & Trim$(WG03.U_MSTCOD) & "' "
        //			//    sQry = sQry & "  AND ISNULL(U_FILD01, '') <> '' AND ISNULL(U_FILD03, 0) <> 0"
        //			//    sQry = sQry & " ORDER BY CODE, LineID"
        //			//    oRecordSet.DoQuery sQry
        //			//    Do Until oRecordSet.EOF
        //			//        For iCol = 1 To 18
        //			//            If oRecordSet.Fields("U_FILD01").Value = WK_G.GONCOD(iCol) Then
        //			//                WG03.U_GONAMT(iCol) = oRecordSet.Fields("U_FILD03").Value
        //			//            End If
        //			//        Next iCol
        //			//        oRecordSet.MoveNext
        //			//    Loop
        //			/// 일할계산 방법 구함(말일기준, 30일기준)
        //			sQry = "SELECT TOP 1 U_INPDAY FROM [@PH_PY107A] WHERE CODE <= '" + oYM + "' ORDER BY CODE DESC";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount > 0) {
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				INPCheck = oRecordSet.Fields.Item(0).Value;
        //				if (Convert.ToDouble(INPCheck) == 2) {
        //					//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //					ENDDAY = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StrDate, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(EndDate, "0000-00-00"))) + 1;
        //					/// 한샘, 케어라인
        //				} else {
        //					ENDDAY = 30;
        //				}
        //			} else {
        //				ENDDAY = 30;
        //			}

        //			/// 2.2) 수습사원 급여 수습율로 구함.
        //			if (Conversion.Val(sRecordset.Fields.Item("U_INPRAT").Value) > 0) {
        //				INPRAT = Conversion.Val(sRecordset.Fields.Item("U_INPRAT").Value) / 100;
        //				/// 율
        //				for (iCol = 1; iCol <= 24; iCol++) {
        //					//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WG03.U_CSUAMT[iCol] * INPRAT, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //				}
        //			}
        //			/// 근태자료관련 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			short WG23_WCHDAY = 0;
        //			short WG23_HGNDAY = 0;
        //			short WG23_CHLDAY = 0;
        //			short WG23_GULDAY = 0;
        //			short WG23_WCHHGA = 0;
        //			short WG23_SNHDAY = 0;
        //			short WG23_JCHDAY = 0;
        //			short WG23_YCHHGA = 0;
        //			short WG23_SNHHGA = 0;
        //			short WG23_JIGCNT = 0;
        //			short WG23_JOTCNT = 0;
        //			double WG23_HYUJUP = 0;
        //			double WG23_NHTTIM = 0;
        //			double WG23_UGBTIM = 0;
        //			double WG23_GNTTIM = 0;
        //			double WG23_JUPTIM = 0;
        //			double WG23_HGNTIM = 0;
        //			double WG23_HYUNHT = 0;
        //			double WG23_JOTTIM = 0;
        //			double WG23_JIGTIM = 0;
        //			double WG23_OICTIM = 0;

        //			sQry = " SELECT T0.*";
        //			sQry = sQry + " FROM [@ZPY230L] T0 INNER JOIN [@ZPY230H] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE   T1.U_GNTYMM = '" + oYM + "'";
        //			sQry = sQry + " AND     T0.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				WG23_CHLDAY = 0;
        //				WG23_HGNDAY = 0;
        //				WG23_GULDAY = 0;
        //				WG23_WCHDAY = 0;
        //				WG23_WCHHGA = 0;
        //				WG23_JCHDAY = 0;
        //				WG23_YCHHGA = 0;
        //				WG23_SNHDAY = 0;
        //				WG23_SNHHGA = 0;
        //				WG23_JIGCNT = 0;
        //				WG23_JOTCNT = 0;
        //				WG23_GNTTIM = 0;
        //				WG23_UGBTIM = 0;
        //				WG23_JUPTIM = 0;
        //				WG23_NHTTIM = 0;
        //				WG23_HGNTIM = 0;
        //				WG23_HYUJUP = 0;
        //				WG23_HYUNHT = 0;
        //				WG23_JIGTIM = 0;
        //				WG23_JOTTIM = 0;
        //				WG23_OICTIM = 0;
        //			} else {
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_CHLDAY = oRecordSet.Fields.Item("U_CHLDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_HGNDAY = oRecordSet.Fields.Item("U_HGNDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_GULDAY = oRecordSet.Fields.Item("U_GULDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_WCHDAY = oRecordSet.Fields.Item("U_WCHDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_WCHHGA = oRecordSet.Fields.Item("U_WCHHGA").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JCHDAY = oRecordSet.Fields.Item("U_JCHDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_YCHHGA = oRecordSet.Fields.Item("U_YCHHGA").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_SNHDAY = oRecordSet.Fields.Item("U_SNHDAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_SNHHGA = oRecordSet.Fields.Item("U_SNHHGA").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JIGCNT = oRecordSet.Fields.Item("U_JIGCNT").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JOTCNT = oRecordSet.Fields.Item("U_JOTCNT").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_GNTTIM = oRecordSet.Fields.Item("U_GNTTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_UGBTIM = oRecordSet.Fields.Item("U_UGBTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JUPTIM = oRecordSet.Fields.Item("U_JUPTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_NHTTIM = oRecordSet.Fields.Item("U_NHTTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_HGNTIM = oRecordSet.Fields.Item("U_HGNTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_HYUJUP = oRecordSet.Fields.Item("U_HYUJUP").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_HYUNHT = oRecordSet.Fields.Item("U_HYUNHT").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JIGTIM = oRecordSet.Fields.Item("U_JIGTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_JOTTIM = oRecordSet.Fields.Item("U_JOTTIM").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG23_OICTIM = oRecordSet.Fields.Item("U_OICTIM").Value;

        //			}
        //			/// 2.3)연봉직,월급직의 당월입퇴사자 일할계산
        //			//// 보통 입퇴사일기준으로 기본급, 고정수당 일할계산함.
        //			if (Strings.Trim(WG03.U_PAYTYP) == "1" | Strings.Trim(WG03.U_PAYTYP) == "2") {
        //				//        Select Case MDC_COMpanyGubun
        //				//Case "CL" '// 케어라인은 관리직도 입퇴사자 입퇴사일기준으로 일할계산하지 않고 월근태를 사용함.
        //				//        Case Else
        //				MDC_SetMod.Term2(ref StrDate, ref EndDate);
        //				if ((WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate) | (WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate)) {
        //					switch (true) {
        //						case (WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate) & (WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate):
        //							/// 당월입퇴사자 기본급 /
        //							MDC_SetMod.Term2(ref WG03.U_INPDAT, ref WG03.U_OUTDAT);
        //							break;
        //						case WG03.U_INPDAT >= StrDate & WG03.U_INPDAT <= EndDate:
        //							/// 당월입사자 기본급 /
        //							MDC_SetMod.Term2(ref WG03.U_INPDAT, ref EndDate);
        //							break;
        //						case WG03.U_OUTDAT >= StrDate & WG03.U_OUTDAT <= EndDate:
        //							/// 당월퇴사자 기본급 /
        //							MDC_SetMod.Term2(ref StrDate, ref WG03.U_OUTDAT);
        //							break;
        //					}
        //					for (iCol = 1; iCol <= 24; iCol++) {
        //						//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WG03.U_CSUAMT[iCol] / ENDDAY * (MDC_Globals.ZPAY_GBL_GNSMON * ENDDAY + MDC_Globals.ZPAY_GBL_GNSDAY), WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //					}
        //				}
        //				//        End Select
        //			}
        //			/// 3) 일급직, 시급직 근태 일할 계산
        //			switch (Strings.Trim(WG03.U_PAYTYP)) {
        //				case "1":
        //					/// 1.연봉직
        //					WK_DAYAMT = sRecordset.Fields.Item("U_STDAMT").Value / 365;
        //					WK_TIMAMT = WK_DAYAMT / 8;
        //					//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WG03.U_STDAMT = MDC_SetMod.IInt(ref WK_TIMAMT, ref 1);
        //					break;
        //				case "2":
        //					/// 2.월급직
        //					WK_DAYAMT = sRecordset.Fields.Item("U_STDAMT").Value / 30;
        //					WK_TIMAMT = WK_DAYAMT / 8;
        //					//한샘
        //					//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WG03.U_STDAMT = MDC_SetMod.IInt(ref WK_TIMAMT, ref 1);
        //					break;
        //				case "3":
        //					/// 3.일급직
        //					//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WK_DAYAMT = sRecordset.Fields.Item("U_STDAMT").Value;
        //					WK_TIMAMT = WK_DAYAMT / 8;
        //					WG03.U_STDAMT = WK_DAYAMT;
        //					break;
        //				case "4":
        //					/// 4.시급직
        //					//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					WK_TIMAMT = sRecordset.Fields.Item("U_STDAMT").Value;
        //					WK_DAYAMT = WK_TIMAMT * 8;
        //					WG03.U_STDAMT = WK_TIMAMT;
        //					break;
        //			}
        //			Imsi_DAYAMT = 0;
        //			Imsi_TIMAMT = 0;
        //			WG03.U_DAYAMT = WK_DAYAMT;

        //			//   '/ 근태관리
        //			/// 근태계산대상인 체크인자만.
        //			if (Strings.Trim(sRecordset.Fields.Item("U_GNTEXE").Value) == "Y") {
        //				for (iCol = 1; iCol <= 24; iCol++) {
        //					/// 생산직만
        //					if (Strings.Trim(sRecordset.Fields.Item("U_BX1SEL").Value) == "Y") {
        //						if (Strings.Trim(WK_C.CSUCOD[iCol]) == "A01" & (Strings.Trim(WG03.U_PAYTYP) == "3" | Strings.Trim(WG03.U_PAYTYP) == "4")) {
        //							//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref (WG23_GNTTIM) * WK_TIMAMT, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //						}
        //						//                    If MDC_COMpanyGubun = "CL" Then
        //						//                        '/ (2007-05-15케어라인)생산직만-장려수당 15일이상 근무시 100%지급, 15일이하이면 지급제외.
        //						//                        If Trim$(WK_C.CSUCOD(iCol)) = "E02" And WG23_CHLDAY < 15 Then
        //						//                            WG03.U_CSUAMT(iCol) = 0
        //						//                        End If
        //						//                        '/ 케어라인: 일할계산(관리직,생산직모두)(결근,병가있으면 자가계발비, 식대보조,교통비 일할계산)
        //						//                        If WG23_GULDAY > 0 Or ((WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) _
        //						//'                                                 Or (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)) Then
        //						//                            If Trim$(WK_C.CSUCOD(iCol)) = "E03" Then '/ 자가계발비
        //						//                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                            ElseIf Trim$(WK_C.CSUCOD(iCol)) = "B01" Then '/ 식대보조
        //						//                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                            ElseIf Trim$(WK_C.CSUCOD(iCol)) = "B02" Then '/ 교통비
        //						//                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                            End If
        //						//                        End If
        //						//                    ElseIf MDC_COMpanyGubun = "SO" Then
        //						//                        '/ 씬터온: 생산장려 결근이 있으면 생산장려지급안함.
        //						//                        If WG23_GULDAY > 0 Or ((WG03.U_OUTDAT >= StrDate And WG03.U_OUTDAT <= EndDate) _
        //						//'                                                 Or (WG03.U_INPDAT >= StrDate And WG03.U_INPDAT <= EndDate)) Then
        //						//                            If Trim$(WK_C.CSUCOD(iCol)) = "E02" Then '/ 생산장려
        //						//                                WG03.U_CSUAMT(iCol) = 0
        //						//                            End If
        //						//                        End If
        //						//
        //						//                    End If
        //					} else {
        //						//                    If MDC_COMpanyGubun = "CL" Then
        //						//                        If Trim$(WG03.U_PAYTYP) = "1" Or Trim$(WG03.U_PAYTYP) = "2" Then
        //						//                            Select Case Trim$(WK_C.CSUCOD(iCol))
        //						//                            Case "E01" '/ (2007-05-15케어라인)관리직만-정근수당 20일이상 100%지급 20일 이하이면 일할계산
        //						//                                 If Trim$(WK_C.CSUCOD(iCol)) = "E01" And WG23_CHLDAY < 20 Then
        //						//                                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                                 End If
        //						//                            Case Else
        //						//                                 WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WG03.U_CSUAMT(iCol) / ENDDAY * WG23_CHLDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                            End Select
        //						//                        End If
        //						//                    End If
        //					}
        //					//// 근태관련 수당들 계산
        //					switch (Strings.Trim(WK_C.CSUCOD[iCol])) {
        //						case "D02":
        //							/// 케어라인 : 연장수당, 한샘 : 야간수당, 씬터온:잔업수당
        //							//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WK_TIMAMT * WG23_JUPTIM * 1.5, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //							break;
        //						case "D04":
        //							/// 특근수당
        //							//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WK_TIMAMT * WG23_HGNTIM * 1.5, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //							break;
        //						case "D06":
        //							/// 특근연장
        //							//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WK_TIMAMT * WG23_HYUJUP * 2, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //							break;
        //						case "D03":
        //							/// 야근수당, 한샘: 심야수당, 씬터온:야간수당
        //							//잔업수당(관리직은고정급, 생산직은 시급계산)
        //							//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = MDC_SetMod.IInt(ref WK_TIMAMT * WG23_NHTTIM * 2, ref 1);
        //							/// 한샘
        //							break;
        //						//                   WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(Imsi_TIMAMT * WG23_NHTTIM * 0.5, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						case "C05":
        //							/// 주휴수당 (씬터온주휴있음)
        //							/// 한샘 반장(월급직)은 주휴수당안나감
        //							if ((Strings.Trim(WG03.U_PAYTYP) == "3" | Strings.Trim(WG03.U_PAYTYP) == "4")) {
        //								//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WK_DAYAMT * WG23_JCHDAY, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //							}
        //							break;
        //						case "D01":
        //							/// 일신: 유급수당, 기본연장  한샘 : 잔업수당 =연장시간
        //							break;
        //						//                        Select Case MDC_COMpanyGubun
        //						//                        Case "CL"
        //						//                                '/ 추가연장: 5.30 별정직 기본연장이 고정급으로 이미들어가있으면 월근태에서 연장계산 제외
        //						//                                If WG03.U_CSUAMT(iCol) = 0 Then
        //						//                                    WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(WK_TIMAMT * WG23_UGBTIM * 1.25, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //						//                                End If
        //						//                        Case "HS"
        //						//                                If Trim$(sRecordset.Fields("U_BX1SEL").Value) = "Y" Then  '/ 생산직만계산
        //						//                                   WG03.U_CSUAMT(iCol) = MDC_SetMod.IInt(WK_TIMAMT * WG23_JUPTIM * 1.5, 1)           '/ 한샘
        //						//                                End If
        //						//                        End Select
        //						case "C02":
        //							/// 월차수당(일신)씬터온월차없음
        //							if ((WG23_WCHDAY - WG23_WCHHGA) > 0) {
        //								//                        If MDC_COMpanyGubun = "IW" Then
        //								//                            WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").Value / 30 * (WG23_WCHDAY - WG23_WCHHGA), WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //								//                        Else
        //								//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref WK_DAYAMT * (WG23_WCHDAY - WG23_WCHHGA), WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //								//                        End If
        //							}
        //							break;
        //						case "C03":
        //							/// 보건수당(일신)씬터온보건없음
        //							if ((WG23_SNHDAY - WG23_SNHHGA) > 0) {
        //								//                            If MDC_COMpanyGubun = "IW" Then
        //								//                                WG03.U_CSUAMT(iCol) = MDC_SetMod.RInt(sRecordset.Fields("U_STDAMT").Value / 30, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)) * (WG23_SNHDAY - WG23_SNHHGA)
        //								//                            Else
        //								//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref sRecordset.Fields.Item("U_STDAMT").Value / 30, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]) * (WG23_SNHDAY - WG23_SNHHGA);
        //								//                            End If
        //							}
        //							break;
        //						case "A01":
        //							break;
        //						//                    '// 케어라인 (2007.05.21 주차수당없이 기본수당에 포함됨.)
        //						//                    If MDC_COMpanyGubun = "CL" Then
        //						//                        If (Trim$(WG03.U_PAYTYP) = "3" Or Trim$(WG03.U_PAYTYP) = "4") Then     '/ 한샘 반장(월급직)은 주휴수당안나감
        //						//                           WG03.U_CSUAMT(iCol) = WG03.U_CSUAMT(iCol) + (MDC_SetMod.RInt(WK_DAYAMT * WG23_JCHDAY, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol)))
        //						//                        End If
        //						//                    End If
        //					}

        //				}
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			pRecordset = null;
        //			return;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			pRecordset = null;
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("PayRoll_Process1 Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //		}
        #endregion

        #region BonusRoll_Process
        //		private string BonusRoll_Process()
        //		{
        //			string functionReturnValue = null;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			short ErrNum = 0;
        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;
        //			int iCol = 0;
        //			int kCol = 0;
        //			string U_INPGBN = null;

        //			string SuSilStr = null;
        //			double Tmp_CSUAMT = 0;
        //			short WK_GNMDAY = 0;
        //			//UPGRADE_WARNING: S_GNMMON 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        //			short[] S_GNMMON = new short[9];
        //			short BNS_RODLEN = 0;
        //			string BNS_ROUNDT = null;
        //			string WK_ENDDAT = null;
        //			string WK_LMTDAT = null;
        //			string BnsResult = null;


        //			ErrNum = 0;
        //			BNS_ROUNDT = "F";
        //			BNS_RODLEN = 10;
        //			WK_ENDDAT = Strings.Trim(oDS_PH_PY111A.GetValue("U_bGNEDAT", 0));
        //			WK_LMTDAT = Strings.Trim(oDS_PH_PY111A.GetValue("U_bEXPDAT", 0));
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			///1. 기본급여사항가져오기
        //			//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WK_STDAMT = sRecordset.Fields.Item("U_BNSAMT").Value;
        //			/// 기본 상여금

        //			//// 1.1. 공식가져오기-수당
        //			sQry = " SELECT T0.U_LINSEQ, T0.U_CSUCOD, T0.U_SILCUN, T0.U_SILCOD, T1.U_BNSLEN, T1.U_BNSRND ";
        //			sQry = sQry + " FROM [@PH_PY106B] T0 INNER JOIN [@PH_PY106A] T1 ON T0.DocEntry = T1.DocEntry ";
        //			sQry = sQry + " WHERE   T1.U_YM = '" + Strings.Trim(U_CSUCOD) + "'";
        //			sQry = sQry + " AND     T1.U_PAYTYP = '" + Strings.Trim(WG03.U_PAYTYP) + "'";
        //			sQry = sQry + " ORDER BY CAST(T0.U_LINSEQ AS INT), T0.U_CSUCOD ";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				ErrNum = 1;
        //				goto Error_Message;
        //			}

        //			//// 2010.04.05 최동권 추가
        //			//// 수당 계산식 설정에서 상여자리수와 끝전처리 세팅값을 가져옴
        //			//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //			if (Information.IsDBNull(oRecordSet.Fields.Item("U_BNSLEN").Value) == false) {
        //				BNS_RODLEN = Conversion.Val(oRecordSet.Fields.Item("U_BNSLEN").Value);
        //			}
        //			//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //			if (Information.IsDBNull(oRecordSet.Fields.Item("U_BNSRND").Value) == false) {
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				BNS_ROUNDT = oRecordSet.Fields.Item("U_BNSRND").Value;
        //			}

        //			while (!(oRecordSet.EoF)) {
        //				iCol = Conversion.Val(oRecordSet.Fields.Item("U_LINSEQ").Value);
        //				if (iCol > 0) {
        //					if (oJOBTYP == "2") {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WK_C.GONSIL[iCol] = oRecordSet.Fields.Item("U_SILCOD").Value;
        //						//// 상여계산식 읽어오기
        //					} else {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WK_C.GONSIL[iCol] = oRecordSet.Fields.Item("U_SILCUN").Value;
        //						//// 급여계산식 읽어오기
        //					}
        //				} else {
        //					if (Strings.Left(oRecordSet.Fields.Item("U_CSUCOD").Value, 1) == "X") {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SuSilStr = oRecordSet.Fields.Item("U_SILCUN").Value;
        //						/// 2.2. 공식안 시스템제공코드값으로 변경
        //						SuSilStr = Change_GOSIL(ref SuSilStr);
        //						switch (Strings.Trim(oRecordSet.Fields.Item("U_CSUCOD").Value)) {
        //							case "X01":
        //								/// 기본일급
        //								X01_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X02":
        //								/// 기본시급
        //								X02_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X03":
        //								/// 통상일급
        //								X03_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //							case "X04":
        //								/// 통상시급
        //								X04_Val = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //								break;
        //						}
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}
        //			WG03.U_DAYAMT = X01_Val;
        //			WG03.U_BASAMT = X02_Val;

        //			//// 2. 공식가지고 계산하기-수당
        //			SuSilStr = "";
        //			WK_BNSAMT = 0;
        //			for (iCol = 1; iCol <= 24; iCol++) {
        //				/// 수당항목이 있는것만
        //				if (!string.IsNullOrEmpty(Strings.Trim(WK_C.CSUCOD[iCol]))) {
        //					if (Strings.Trim(WK_C.BNSUSE[iCol]) == "Y") {
        //						/// 공식이있으면 공식계산
        //						if (!string.IsNullOrEmpty(Strings.Trim(WK_C.GONSIL[iCol]))) {
        //							/// 2.1. 공식-계산결과값가져오는거면..
        //							SuSilStr = WK_C.GONSIL[iCol];
        //							/// 계산된값일 경우
        //							for (kCol = 1; kCol <= 24; kCol++) {
        //								SuSilStr = Strings.Replace(SuSilStr, "#" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(kCol, "00"), Convert.ToString(WG03.U_CSUAMT[kCol]));
        //							}
        //							/// 2.2. 공식안 시스템제공코드값으로 변경
        //							SuSilStr = Change_GOSIL(ref SuSilStr);
        //							/// 2.3. 공식계산하기
        //							Tmp_CSUAMT = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //							/// 2.4. 정답가져오기

        //							///2010.04.05 최동권 수정 (상여끝전 처리값을 수당계산식 설정에서 가져옴)
        //							//                    If Trim$(WK_C.CSUCOD(iCol)) = "A04" Or Trim$(WK_C.CSUCOD(iCol)) = "A01" Then
        //							//                        BNS_RODLEN = WK_C.RODLEN(iCol)
        //							//                        BNS_ROUNDT = WK_C.ROUNDT(iCol)
        //							//                    End If
        //							//// 급상여:상여금수당만
        //							if (oJOBTYP == "3" & Strings.Trim(WK_C.CSUCOD[iCol]) == "A04") {
        //								//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref Tmp_CSUAMT, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //								WK_BNSAMT = WG03.U_CSUAMT[iCol];
        //							//// 상여:상여모두
        //							} else if (oJOBTYP == "2") {
        //								//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_CSUAMT[iCol] = MDC_SetMod.RInt(ref Tmp_CSUAMT, WK_C.RODLEN[iCol], WK_C.ROUNDT[iCol]);
        //								if (Strings.Trim(WK_C.CSUCOD[iCol]) == "A01" & WG03.U_CSUAMT[iCol] == 0)
        //									WG03.U_CSUAMT[iCol] = WK_STDAMT;
        //								WK_BNSAMT = WK_BNSAMT + WG03.U_CSUAMT[iCol];
        //							}
        //						}
        //					}
        //				}
        //			}

        //			//// 2010.04.05 최동권 추가
        //			//// 기간급여 평균으로 상여 계산(TNC 이슈사항)
        //			if (oBNSCAL == "2" | oBNSCAL == "3") {
        //				BnsResult = BonusRoll_Process2();
        //				if (!string.IsNullOrEmpty(BnsResult)) {
        //					ErrNum = 2;
        //					goto Error_Message;
        //				}
        //			}


        //			//        SuSilStr = ""
        //			//        If oRecordSet.RecordCount > 0 Then
        //			//            SuSilStr = oRecordSet.Fields("U_SILCUN").Value   '/ 계산된값일 경우
        //			//            '/ 2.2. 공식안 시스템제공코드값으로 변경
        //			//            SuSilStr = Change_GOSIL(SuSilStr)   '/ 월총일수
        //			//            If Trim$(SuSilStr) <> "" Then
        //			//                WK_BNSAMT = Get_ReAmt(oYM, WG03.U_MSTCOD, SuSilStr)
        //			//            Else
        //			//                WK_BNSAMT = 0
        //			//            End If
        //			//        Else
        //			//            WK_BNSAMT = sRecordSet.Fields("U_BNSAMT").Value
        //			//        End If
        //			//    End If
        //			///3.평균임금관련(급여세액기간)/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			//    'sQry = "SELECT SUM(CASE WHEN U_JOBTYP ='2' THEN 0 ELSE 1 END) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
        //			//    '200907.22 수정 같은월 급여 2개일경우 2로 잡힘.
        //			//    sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
        //			//    sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
        //			//    sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
        //			//    sQry = sQry & "                    AND U_YM >= '" & Mid$(oSTRTAX, 1, 6) & "'"
        //			//    If Trim$(oENDTAX) < oYM Then  '// 세액종료연월이 귀속연월보다 적을경우
        //			//        sQry = sQry & "                AND U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
        //			//    Else                                   '// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
        //			//        sQry = sQry & "                AND (U_YM <= '" & Mid$(oENDTAX, 1, 6) & "'"
        //			//        sQry = sQry & "                AND NOT (U_YM = '" & oYM & "' AND (U_JOBTYP = '" & oJOBTYP & "' AND U_JOBGBN = '" & oJOBGBN & "')))"
        //			//    End If
        //			//    oRecordSet.DoQuery sQry
        //			//    If oRecordSet.RecordCount <> 0 Then
        //			//       WG03.U_AVRPAY = IIf(IsNull(oRecordSet.Fields("TOTPAY").Value), 0, oRecordSet.Fields("TOTPAY").Value)
        //			//       WG03.U_NABTAX = IIf(IsNull(oRecordSet.Fields("TOTGAB").Value), 0, oRecordSet.Fields("TOTGAB").Value)
        //			//'       WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
        //			//    End If
        //			//    '/ 상여나가는 달 세액기간
        //			//    TermCHK = False
        //			//    If (oJOBTYP = "2" Or oJOBTYP = "3") Then
        //			//        sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,"
        //			//        sQry = sQry & "   SUM(U_GONG05) AS TOTGBH"
        //			//        sQry = sQry & " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" & Trim$(WG03.U_MSTCOD) & "'"
        //			//        sQry = sQry & "                      AND U_YM >= '" & Mid$(oSTRTAX, 1, 6) & "'"
        //			//        sQry = sQry & "                      AND (U_YM < '" & Mid$(oENDTAX, 1, 6) & "'"
        //			//        sQry = sQry & "                      AND U_JOBTYP <> '2'"
        //			//        sQry = sQry & "                      OR (U_YM = '" & Mid$(oENDTAX, 1, 6) & "' AND (U_JOBTYP = '1' OR U_JOBTYP='3')))"
        //			//        oRecordSet.DoQuery sQry
        //			//        If oRecordSet.RecordCount <> 0 Then
        //			//           WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
        //			//        End If
        //			//        If oJOBTYP = "2" And WG03.U_TAXTRM = 0 Then
        //			//            TermCHK = True  '/ 상여만일경우 세액기간이 0일경우 확인메세지출력
        //			//        End If
        //			//    End If

        //			////*********************************************************************************************************************/
        //			/// 2010.02.23  함미경 수정 (관련업체:대우IS등 )
        //			/// 수정 상여가 먼저나가는 업체 세액계산시 급여가져오는 세액기간과 상여가져오는 세액기간 별도 구분
        //			/// 위 세액기간 2개월동안 급여1개월 상여2개월분일경우 평균임금/1개월로 계산 세액높게 책정, 생성할 상여달 급여미지급일경우 급여세액기간 전전월~전월,상여 전월~당월
        //			/// 급여계산기준으로 개월수 계산하도록 변경함
        //			////*********************************************************************************************************************/
        //			/// 상여기간계산
        //			sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,";
        //			sQry = sQry + "   SUM(U_GONG05) AS TOTGBH";
        //			sQry = sQry + " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			sQry = sQry + "                    AND U_YM >= '" + Strings.Mid(oSTRBNS, 1, 6) + "'";
        //			//// 세액종료연월이 귀속연월보다 적을경우
        //			if (Strings.Trim(oENDBNS) < oYM) {
        //				sQry = sQry + "                AND U_YM <= '" + Strings.Mid(oENDBNS, 1, 6) + "'";
        //			//// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
        //			} else {
        //				sQry = sQry + "                AND (U_YM <= '" + Strings.Mid(oENDBNS, 1, 6) + "'";
        //				sQry = sQry + "                AND NOT (U_YM = '" + oYM + "' AND (U_JOBTYP = '" + oJOBTYP + "' AND U_JOBGBN = '" + oJOBGBN + "')))";
        //			}
        //			sQry = sQry + "                    AND U_JOBTYP = '2'";
        //			/// 상여만
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount != 0) {
        //				//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //				WG03.U_AVRPAY = (Information.IsDBNull(oRecordSet.Fields.Item("TOTPAY").Value) ? 0 : oRecordSet.Fields.Item("TOTPAY").Value);
        //				//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //				WG03.U_NABTAX = (Information.IsDBNull(oRecordSet.Fields.Item("TOTGAB").Value) ? 0 : oRecordSet.Fields.Item("TOTGAB").Value);
        //				//       WG03.U_TAXTRM = IIf(IsNull(oRecordSet.Fields("TOTTRM").Value), 0, oRecordSet.Fields("TOTTRM").Value)
        //			}
        //			/// 상여나가는 달 세액기간
        //			TermCHK = false;
        //			if ((oJOBTYP == "2" | oJOBTYP == "3")) {
        //				sQry = "SELECT COUNT(DISTINCT U_YM) AS TOTTRM, SUM(U_GWASEE) AS TOTPAY, SUM(U_GONG01) AS TOTGAB,";
        //				sQry = sQry + "   SUM(U_GONG05) AS TOTGBH";
        //				sQry = sQry + " FROM  [@PH_PY111A] WHERE U_MSTCOD  = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //				sQry = sQry + "                      AND U_YM >= '" + Strings.Mid(oSTRTAX, 1, 6) + "'";
        //				//// 세액종료연월이 귀속연월보다 적을경우
        //				if (Strings.Trim(oENDTAX) < oYM) {
        //					sQry = sQry + "                AND U_YM <= '" + Strings.Mid(oENDTAX, 1, 6) + "'";
        //				//// 세액종료연월보다 크거나 같으면 해당 계산중인 귀속연월,지급종류,지급구분은 제외
        //				} else {
        //					sQry = sQry + "                AND (U_YM <= '" + Strings.Mid(oENDTAX, 1, 6) + "'";
        //					sQry = sQry + "                AND NOT (U_YM = '" + oYM + "' AND (U_JOBTYP = '" + oJOBTYP + "' AND U_JOBGBN = '" + oJOBGBN + "')))";
        //				}
        //				sQry = sQry + "                    AND U_JOBTYP <> '2'";
        //				/// 상여만
        //				oRecordSet.DoQuery(sQry);
        //				if (oRecordSet.RecordCount != 0) {
        //					//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //					WG03.U_AVRPAY = WG03.U_AVRPAY + (Information.IsDBNull(oRecordSet.Fields.Item("TOTPAY").Value) ? 0 : oRecordSet.Fields.Item("TOTPAY").Value);
        //					//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //					WG03.U_NABTAX = WG03.U_NABTAX + (Information.IsDBNull(oRecordSet.Fields.Item("TOTGAB").Value) ? 0 : oRecordSet.Fields.Item("TOTGAB").Value);
        //					//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        //					WG03.U_TAXTRM = (Information.IsDBNull(oRecordSet.Fields.Item("TOTTRM").Value) ? 0 : oRecordSet.Fields.Item("TOTTRM").Value);
        //				}
        //				if (oJOBTYP == "2" & WG03.U_TAXTRM == 0) {
        //					TermCHK = true;
        //					/// 상여만일경우 세액기간이 0일경우 확인메세지출력
        //				}
        //			}

        //			///4. 근속년월일 계산 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			S_GNMMON[1] = 0;
        //			S_GNMMON[2] = 0;
        //			S_GNMMON[3] = 0;
        //			S_GNMMON[4] = 0;
        //			S_GNMMON[5] = 0;
        //			S_GNMMON[6] = 0;
        //			S_GNMMON[7] = 0;
        //			S_GNMMON[8] = 0;
        //			WK_GNMDAY = 0;
        //			MDC_SetMod.Term2(ref WG03.U_INPDAT, ref WK_ENDDAT);
        //			WG03.U_GNSYER = MDC_Globals.ZPAY_GBL_GNSYER;
        //			WG03.U_GNSMON = MDC_Globals.ZPAY_GBL_GNSMON;

        //			//UPGRADE_WARNING: DateDiff 동작이 다를 수 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        //			WK_GNMDAY = DateAndTime.DateDiff(Microsoft.VisualBasic.DateInterval.Day, Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00")), Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WK_ENDDAT, "0000-00-00"))) + 1;
        //			/// 적용1
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP1GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[1] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[1] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용2
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP2GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[2] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[2] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용3
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP3GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[3] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[3] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용4
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP4GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[4] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[4] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용5
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP5GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[5] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[5] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용6
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP6GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[6] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[6] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용7
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP7GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[7] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[7] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			/// 적용8
        //			switch (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP8GBN", 0))) {
        //				case "3":
        //				case "4":
        //					S_GNMMON[8] = WK_GNMDAY;
        //					/// 일수
        //					break;
        //				default:
        //					S_GNMMON[8] = MDC_Globals.ZPAY_GBL_GNMMON;
        //					/// 개월수
        //					break;
        //			}
        //			///5. 상여금계산  /

        //			if (oJOBTYP == "2") {
        //				WG03.U_TAXTRM = (WG03.U_TAXTRM < 1 ? 1 : WG03.U_TAXTRM);
        //			}
        //			WK_APPBNS = 0;
        //			WG03.U_BNSRAT = Convert.ToDouble(Strings.Trim(oDS_PH_PY111A.GetValue("U_bBNSRAT", 0)));

        //			///*******입사구분 경력으로 들어왔음 상관없이 개월수 상관없이 적용율 100% 지급함 (신입인경우만 상여율 적용혀)
        //			//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			U_INPGBN = MDC_SetMod.Get_ReData(ref "U_RelCd", ref "U_Minor", ref "[@PS_HR200L]", ref "'" + sRecordset.Fields.Item("U_INPGBN").Value + "'", ref " AND Code =N'P133'");
        //			//// 3-경력자이면 100적용율
        //			if (Strings.Trim(U_INPGBN) == "3") {
        //				WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0));
        //				/// 경력직이면 맨 상위적용율 적용되도록.
        //				WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //				WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //			} else {
        //				switch (true) {
        //					/// 적용1: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) != "4") & S_GNMMON[1] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH1", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //						break;
        //					/// 적용1: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP1GBN", 0)) == "4") & S_GNMMON[1] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH1", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT1", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT1", 0));
        //						break;
        //					/// 적용2: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) != "4") & S_GNMMON[2] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH2", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT2", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0));
        //						break;
        //					/// 적용2: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP2GBN", 0)) == "4") & S_GNMMON[2] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH2", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT2", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT2", 0));
        //						break;
        //					/// 적용3: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) != "4") & S_GNMMON[3] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH3", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT3", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0));
        //						break;
        //					/// 적용3: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP3GBN", 0)) == "4") & S_GNMMON[3] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH3", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT3", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT3", 0));
        //						break;
        //					/// 적용4: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) != "4") & S_GNMMON[4] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH4", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT4", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0));
        //						break;
        //					/// 적용4: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP4GBN", 0)) == "4") & S_GNMMON[4] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH4", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT4", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT4", 0));
        //						break;
        //					/// 적용5: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) != "4") & S_GNMMON[5] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH5", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT5", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0));
        //						break;
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP5GBN", 0)) == "4") & S_GNMMON[5] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH5", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT5", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT5", 0));
        //						break;
        //					/// 적용6: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) != "4") & S_GNMMON[6] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH6", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT6", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0));
        //						break;
        //					/// 적용6: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP6GBN", 0)) == "4") & S_GNMMON[6] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH6", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT6", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT6", 0));
        //						break;
        //					/// 적용7: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) != "4") & S_GNMMON[7] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH7", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT7", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0));
        //						break;
        //					/// 적용7: 미만
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP7GBN", 0)) == "4") & S_GNMMON[7] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH7", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT7", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT7", 0));
        //						break;
        //					/// 적용8: 이상
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) != "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) != "4") & S_GNMMON[8] >= Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH8", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT8", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0));
        //						break;
        //					case (Strings.Trim(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) == "2" | Strings.Trim(oDS_PH_PY111A.GetValue("U_AP8GBN", 0)) == "4") & S_GNMMON[8] < Conversion.Val(oDS_PH_PY111A.GetValue("U_bMONTH8", 0)):
        //						WG03.U_APPRAT = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPRAT8", 0));
        //						WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT * (WG03.U_BNSRAT / 100) + Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0));
        //						WK_APPBNS = Conversion.Val(oDS_PH_PY111A.GetValue("U_bAPPAMT8", 0));
        //						break;
        //				}
        //			}

        //			//// 2010.04.07 최동권 추가
        //			//// 개인별 상여율 적용(TNC 이슈사항)
        //			BnsResult = BonusRate_Process();
        //			if (!string.IsNullOrEmpty(BnsResult)) {
        //				ErrNum = 2;
        //				goto Error_Message;
        //			}

        //			//WG03.U_BONUSS = MDC_SetMod.IInt(WG03.U_BONUSS, 10)
        //			//MDC_SetMod.RInt(Tmp_CSUAMT, WK_C.RODLEN(iCol), WK_C.ROUNDT(iCol))
        //			if (string.IsNullOrEmpty(BNS_ROUNDT))
        //				BNS_ROUNDT = "F";
        //			if (BNS_RODLEN == 0)
        //				BNS_RODLEN = 10;
        //			//UPGRADE_WARNING: MDC_SetMod.RInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_BONUSS = MDC_SetMod.RInt(ref WG03.U_BONUSS, BNS_RODLEN, BNS_ROUNDT);
        //			/// 당월퇴사자처리(당월만근시 지급, 중도퇴사자 제외)
        //			if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_OUTDAT))) {
        //				if (WK_LMTDAT > WG03.U_OUTDAT) {
        //					WG03.U_BONUSS = 0;
        //				}
        //			}

        //			//If oJOBTYP = "2" Then    '/ 상여만이면
        //			//    WG03.U_CSUAMT(1) = WG03.U_BONUSS
        //			//Else
        //			//    WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS
        //			//End If

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = "";
        //			return functionReturnValue;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			if (ErrNum == 1) {
        //				functionReturnValue = "수당계산식에 일치하는 자료가 없습니다.";
        //			} else if (ErrNum == 2) {
        //				functionReturnValue = BnsResult;
        //			} else {
        //				functionReturnValue = Err().Number + Strings.Space(10) + Err().Description;
        //			}
        //			return functionReturnValue;
        //		}
        #endregion

        #region BonusRoll_Process2
        ////---------------------------------------------------------------------------------------
        //// Procedure : BonusRoll_Process2
        //// Author    : Choi Dong Kwon
        //// Date      : 2010-04-06
        //// Purpose   : 계산식이 아닌 방법으로 상여계산시 사용
        ////---------------------------------------------------------------------------------------
        //		private string BonusRoll_Process2()
        //		{
        //			string functionReturnValue = null;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			short iCol = 0;
        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			if (oBNSCAL == "2" | oBNSCAL == "3") {
        //				//// Procedure Result
        //				//// U_TOTPAY : 급여 총액
        //				//// U_PAYCNT : 급여 기간 월수
        //				//// U_PAYTRM : 급여 실지급 월수
        //				//// U_AVRPAY : 평균 급여
        //				sQry = "EXEC PH_PY111_BONUSS '" + Strings.Trim(WG03.U_MSTCOD) + "', '" + oSTRTAX + "', '" + oENDTAX + "', '" + oSTRBNS + "', '" + oENDBNS + "', '" + oBNSCAL + "'";
        //				oRecordSet.DoQuery(sQry);
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WK_BNSAMT = oRecordSet.Fields.Item("U_AVRPAY").Value;
        //				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				WG03.U_BONUSS = oRecordSet.Fields.Item("U_AVRPAY").Value;
        //				if (oRecordSet.Fields.Item("U_AVRPAY").Value <= 0) {
        //					WK_BNSAMT = 0;
        //					WG03.U_BONUSS = 0;
        //				}
        //				//// 수당항목 중에 상여금(A04)이 있을 경우 상여금 금액에 넣을 것
        //				for (iCol = 1; iCol <= 24; iCol++) {
        //					if (WK_C.CSUCOD[iCol] == "A04") {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WG03.U_CSUAMT[iCol] = oRecordSet.Fields.Item("U_AVRPAY").Value;
        //					}
        //				}
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = "";
        //			return functionReturnValue;
        //			Error_Message:

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = Err().Description;
        //			return functionReturnValue;

        //		}
        #endregion

        #region BonusRate_Process
        ////---------------------------------------------------------------------------------------
        //// Procedure : BonusRate_Process
        //// Author    : Choi Dong Kwon
        //// Date      : 2010-04-06
        //// Purpose   : 개인별 상여율 계산
        ////---------------------------------------------------------------------------------------
        //		private string BonusRate_Process()
        //		{
        //			string functionReturnValue = null;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//// 개인별 상여율이 등록되어 있는지 체크
        //			sQry = "       SELECT T1.U_BNSRAT, T1.U_BNSAMT ";
        //			sQry = sQry + "FROM   [@ZPY315H] T0 ";
        //			sQry = sQry + "       INNER JOIN [@ZPY315L] T1 ON T0.DocEntry = T1.DocEntry ";
        //			sQry = sQry + "WHERE  T0.U_YM = '" + oYM + "' ";
        //			sQry = sQry + "AND    T0.U_JOBTYP = '" + oJOBTYP + "' ";
        //			sQry = sQry + "AND    T0.U_JOBGBN = '" + oJOBGBN + "' ";
        //			sQry = sQry + "AND    T1.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "' ";
        //			sQry = sQry + "AND   (T1.U_BNSRAT <> 0 OR T1.U_BNSAMT <> 0) ";

        //			oRecordSet.DoQuery(sQry);

        //			//// 개인별 상여율이 등록되어 있지 않은 경우 Skip
        //			if (oRecordSet.RecordCount > 0) {
        //				//        WG03.U_BNSRAT = Val(oRecordSet.Fields("U_BNSRAT").Value)
        //				WG03.U_APPRAT = Conversion.Val(oRecordSet.Fields.Item("U_BNSRAT").Value);
        //				WG03.U_BONUSS = (WG03.U_APPRAT / 100) * WK_BNSAMT + Conversion.Val(oRecordSet.Fields.Item("U_BNSAMT").Value);
        //				WK_APPBNS = WK_BNSAMT;
        //			}
        //			functionReturnValue = "";

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			Error_Message:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = Err().Description;
        //			return functionReturnValue;
        //		}
        #endregion

        #region ChangeItem_Process
        //		private void ChangeItem_Process()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;
        //			short iCol = 0;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			/// 변동사항 조회
        //			sQry = " SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT";
        //			sQry = sQry + " FROM [@ZPY302L] T0 INNER JOIN [@ZPY302H] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE T1.U_YM = '" + oYM + "'";
        //			sQry = sQry + " AND T1.U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND T1.U_JOBGBN = '" + oJOBGBN + "'";
        //			//   sQry = sQry & " AND T1.U_JOBTRG = '" & oJOBTRG & "'"
        //			sQry = sQry + " AND T0.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			sQry = sQry + " ORDER BY T0.U_CSUTYP, T0.U_CSUCOD";
        //			oRecordSet.DoQuery(sQry);
        //			while (!(oRecordSet.EoF)) {
        //				/// 수당항목
        //				if (oRecordSet.Fields.Item("U_CSUTYP").Value == "1") {
        //					for (iCol = 1; iCol <= 24; iCol++) {
        //						if (oRecordSet.Fields.Item("U_CSUCOD").Value == WK_C.CSUCOD[iCol]) {
        //							WG03.U_CSUAMT[iCol] = WG03.U_CSUAMT[iCol] + oRecordSet.Fields.Item("U_CSUAMT").Value;
        //						}
        //					}
        //				/// 공제항목
        //				} else if (oRecordSet.Fields.Item("U_CSUTYP").Value == "2") {
        //					for (iCol = 1; iCol <= 18; iCol++) {
        //						if (oRecordSet.Fields.Item("U_CSUCOD").Value == WK_G.GONCOD[iCol]) {
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + oRecordSet.Fields.Item("U_CSUAMT").Value;
        //						}
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}
        //			/// 건강보험 정산 조회
        //			sQry = " SELECT  SUM(ISNULL(T0.U_CHAMED,0)) AS U_MEDJSN";
        //			sQry = sQry + " FROM [@ZPY311L] T0 INNER JOIN [@ZPY311H] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE T1.U_YM = '" + oYM + "'";
        //			sQry = sQry + " AND T1.U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND T1.U_JOBGBN = '" + oJOBGBN + "'";
        //			sQry = sQry + " AND T1.U_JOBTRG = '" + oJOBTRG + "'";
        //			sQry = sQry + " AND T0.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			oRecordSet.DoQuery(sQry);
        //			while (!(oRecordSet.EoF)) {
        //				for (iCol = 1; iCol <= 18; iCol++) {
        //					if ((G08_CHK == false & WK_G.GONCOD[iCol] == "G04") | (G08_CHK == true & WK_G.GONCOD[iCol] == "G08")) {
        //						WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + oRecordSet.Fields.Item("U_MEDJSN").Value;
        //						WG03.U_TOTGON = WG03.U_TOTGON + oRecordSet.Fields.Item("U_MEDJSN").Value;
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}

        //			/// 급상여변동자료(기간) -변동금액으로 대체/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			sQry = "SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT, T0.U_MSTCOD";
        //			sQry = sQry + " FROM [@ZPY312L] T0 INNER JOIN [@ZPY312H] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE T1.U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND   T1.U_JOBGBN = '" + oJOBGBN + "'";
        //			sQry = sQry + " AND   T0.U_CSUTYP = '1'";
        //			//수당항목만
        //			sQry = sQry + " AND   T0.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			sQry = sQry + " AND   " + "'" + oYM + "' BETWEEN T0.U_STRYMM AND T0.U_ENDYMM";
        //			sQry = sQry + " ORDER BY T0.U_MSTCOD, T1.U_YM DESC, T0.U_STRYMM DESC";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount > 0) {
        //				if (oJOBTYP != "3") {
        //					WG03.U_BONUSS = 0;
        //				}
        //				while (!(oRecordSet.EoF)) {
        //					for (iCol = 1; iCol <= 24; iCol++) {
        //						if (oRecordSet.Fields.Item("U_CSUCOD").Value == WK_C.CSUCOD[iCol]) {
        //							//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_CSUAMT[iCol] = oRecordSet.Fields.Item("U_CSUAMT").Value;
        //							if (oJOBTYP == "3" & Strings.Trim(WK_C.CSUCOD[iCol]) == "A04") {
        //								//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								WG03.U_BONUSS = oRecordSet.Fields.Item("U_CSUAMT").Value;
        //								WG03.U_APPRAT = Convert.ToDouble("100");
        //								// WG03.U_AVRPAY = 0    '/ 2011.10.05 일성신약 변동자료기간별로 상여데이터올려도 평균임금반영하여 계산되도록
        //								// WG03.U_NABTAX = 1
        //								WG03.U_BNSRAT = Convert.ToDouble("100");
        //								WK_APPBNS = WG03.U_BONUSS;
        //							} else if (oJOBTYP == "2" & Strings.Trim(WK_C.BNSUSE[iCol]) == "Y") {
        //								WG03.U_BONUSS = WG03.U_BONUSS + oRecordSet.Fields.Item("U_CSUAMT").Value;
        //								WG03.U_APPRAT = Convert.ToDouble("100");
        //								// WG03.U_AVRPAY = 0    '/ 2011.10.05 일성신약 변동자료기간별로 상여데이터올려도 평균임금반영하여 계산되도록
        //								// WG03.U_NABTAX = 1
        //								WG03.U_BNSRAT = Convert.ToDouble("100");
        //								WK_APPBNS = WG03.U_BONUSS;
        //							}
        //						}
        //					}
        //					oRecordSet.MoveNext();
        //				}
        //			}
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("PH_PY111_Save Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //		}
        #endregion

        #region JeongSan_Process
        //		private void JeongSan_Process()
        //		{
        //			short iCol = 0;

        //			/// 소득세,주민세항목
        //			for (iCol = 1; iCol <= 18; iCol++) {
        //				//If (G06_CHK = False And WK_G.GONCOD(iCol) = "G01") Or (G06_CHK = True And WK_G.GONCOD(iCol) = "G06") Then
        //				if ((G06_CHK == true & WK_G.GONCOD[iCol] == "G06")) {
        //					WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + Conversion.Val(sRecordset.Fields.Item("U_CHAGAB").Value);
        //					//ElseIf (G07_CHK = True And WK_G.GONCOD(iCol) = "G07") Or (G07_CHK = False And WK_G.GONCOD(iCol) = "G02") Then
        //				} else if ((G07_CHK == true & WK_G.GONCOD[iCol] == "G07")) {
        //					WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + Conversion.Val(sRecordset.Fields.Item("U_CHAJUM").Value);
        //				/// 농특세
        //				} else if ((G91_CHK == true & WK_G.GONCOD[iCol] == "G91")) {
        //					WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + Conversion.Val(sRecordset.Fields.Item("U_CHANON").Value);
        //				}
        //			}
        //			/// 총공제금액
        //			WG03.U_TOTGON = WG03.U_TOTGON + Conversion.Val(sRecordset.Fields.Item("U_CHAGAB").Value) + Conversion.Val(sRecordset.Fields.Item("U_CHAJUM").Value) + Conversion.Val(sRecordset.Fields.Item("U_CHANON").Value);
        //			/// 실지급액
        //			//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_SILJIG = MDC_SetMod.IInt(ref WG03.U_TOTPAY - WG03.U_TOTGON, ref 1);

        //			return;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("JeongSan_Process Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //		}
        #endregion

        #region PAY_Tax
        //		private void PAY_Tax()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;
        //			short iCol = 0;
        //			double MEDAMT = 0;
        //			double JUMINN = 0;
        //			double GABGUN = 0;
        //			double GBHAMT = 0;
        //			double KUKAMT = 0;
        //			double BTAX06 = 0;
        //			double BTAX04 = 0;
        //			double BTAX02 = 0;
        //			double GWASEE = 0;
        //			double INCOME = 0;
        //			double BTAX01 = 0;
        //			double BTAX03 = 0;
        //			double BTAX05 = 0;
        //			double BTAX07 = 0;
        //			double GITBTX = 0;
        //			double CHLCSU = 0;
        //			double CARCSU = 0;
        //			double BT1AMT = 0;
        //			double MONPAY = 0;
        //			double FODCSU = 0;
        //			double FRNCSU = 0;
        //			double YGUCSU = 0;
        //			double GITBTX_N = 0;
        //			double TOTGBH = 0;
        //			double BNSCSU = 0;
        //			double BT1KUM = 0;
        //			double GBHRAT = 0;
        //			string SuSilStr = null;
        //			double Tmp_CSUAMT = 0;
        //			string GBHEND = null;
        //			string KUKEND = null;
        //			string KUKSTR = null;
        //			string GBHSTR = null;
        //			double WK_MEDAMT = 0;
        //			string U_NJYYMM = null;
        //			string MEDCHK = null;
        //			double U_NJCRAT = 0;
        //			double U_NJYRAT = 0;
        //			double NJCMED = 0;
        //			int kCol = 0;
        //			double WK_GBHAMT = 0;
        //			//UPGRADE_WARNING: WK_YBTXAM 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        //			double[] WK_YBTXAM = new double[8];
        //			bool WK_BNSCHK = false;

        //			short Dw_Ilsu = 0;

        //			GABGUN = 0;
        //			JUMINN = 0;
        //			GBHAMT = 0;
        //			MEDAMT = 0;
        //			KUKAMT = 0;
        //			INCOME = 0;
        //			GWASEE = 0;
        //			BTAX01 = 0;
        //			BTAX02 = 0;
        //			BTAX03 = 0;
        //			BTAX04 = 0;
        //			BTAX05 = 0;
        //			BTAX06 = 0;
        //			BTAX07 = 0;
        //			MONPAY = 0;
        //			BT1AMT = 0;
        //			FODCSU = 0;
        //			CARCSU = 0;
        //			TOTGBH = 0;
        //			CHLCSU = 0;
        //			YGUCSU = 0;
        //			BNSCSU = 0;
        //			GITBTX = 0;
        //			GITBTX_N = 0;
        //			U_NJYYMM = Convert.ToString(0);
        //			U_NJYRAT = 0;
        //			U_NJCRAT = 0;
        //			NJCMED = 0;

        //			Dw_Ilsu = 0;

        //			///
        //			WK_BNSCHK = false;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			/// 급여총액 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			/// 월정체크, 한도액, 생산비과대상자,부양가족공제,고용보험계산자
        //			for (iCol = 1; iCol <= 24; iCol++) {
        //				if (oJOBTYP == "3" & WK_C.CSUCOD[iCol] == "A04") {
        //					if (oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "N") {
        //						WG03.U_CSUAMT[iCol] = WG03.U_BONUSS;
        //						BNSCSU = BNSCSU + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //					} else {
        //						WG03.U_BONUSS = WG03.U_CSUAMT[iCol];
        //					}
        //				}
        //				if (string.IsNullOrEmpty(Strings.Trim(WK_C.GWATYP[iCol])))
        //					WK_C.GWATYP[iCol] = "1";
        //				switch (WK_C.GWATYP[iCol]) {
        //					case "1":
        //						/// 과세
        //						if (WK_C.CSUKUM[iCol] > 0 & WG03.U_CSUAMT[iCol] > WK_C.CSUKUM[iCol]) {
        //							WG03.U_CSUAMT[iCol] = Conversion.Val(Convert.ToString(WK_C.CSUKUM[iCol]));
        //						}
        //						break;
        //					case "2":
        //						/// 식대보조
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							FODCSU = 0;
        //						} else {
        //							FODCSU = Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & FODCSU > WK_C.CSUKUM[iCol]) {
        //							BTAX02 = BTAX02 + WK_C.CSUKUM[iCol];
        //						} else {
        //							BTAX02 = BTAX02 + FODCSU;
        //						}
        //						break;
        //					case "3":
        //						/// 차량보조
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							CARCSU = 0;
        //						} else {
        //							CARCSU = Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & CARCSU > WK_C.CSUKUM[iCol]) {
        //							BTAX02 = BTAX02 + WK_C.CSUKUM[iCol];
        //						} else {
        //							BTAX02 = BTAX02 + CARCSU;
        //						}
        //						break;
        //					case "4":
        //						/// 생산비과
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							BT1AMT = BT1AMT + 0;
        //						} else {
        //							BT1AMT = BT1AMT + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}

        //						if (BT1KUM < WK_C.CSUKUM[iCol])
        //							BT1KUM = WK_C.CSUKUM[iCol];
        //						//비과세4-마지막줄의 한도액만
        //						break;
        //					case "5":
        //						/// 국외비과세
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							FRNCSU = FRNCSU + 0;
        //						} else {
        //							FRNCSU = FRNCSU + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & FRNCSU > WK_C.CSUKUM[iCol]) {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "M01":
        //									WG03.U_BTXM01 = WG03.U_BTXM01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "M02":
        //									WG03.U_BTXM02 = WG03.U_BTXM02 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "M03":
        //									WG03.U_BTXM03 = WG03.U_BTXM03 + WK_C.CSUKUM[iCol];
        //									break;
        //								default:
        //									BTAX03 = BTAX03 + WK_C.CSUKUM[iCol];
        //									break;
        //							}
        //						} else {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "M01":
        //									WG03.U_BTXM01 = WG03.U_BTXM01 + FRNCSU;
        //									break;
        //								case "M02":
        //									WG03.U_BTXM02 = WG03.U_BTXM02 + FRNCSU;
        //									break;
        //								case "M03":
        //									WG03.U_BTXM03 = WG03.U_BTXM03 + FRNCSU;
        //									break;
        //								default:
        //									BTAX03 = BTAX03 + FRNCSU;
        //									break;
        //							}
        //						}
        //						break;
        //					case "6":
        //						/// 연구수당
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							YGUCSU = YGUCSU + 0;
        //						} else {
        //							YGUCSU = YGUCSU + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & YGUCSU > WK_C.CSUKUM[iCol]) {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "H06":
        //									WG03.U_BTXH06 = WG03.U_BTXH06 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H07":
        //									WG03.U_BTXH07 = WG03.U_BTXH07 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H08":
        //									WG03.U_BTXH08 = WG03.U_BTXH08 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H09":
        //									WG03.U_BTXH09 = WG03.U_BTXH09 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H10":
        //									WG03.U_BTXH10 = WG03.U_BTXH10 + WK_C.CSUKUM[iCol];
        //									break;
        //								default:
        //									BTAX05 = BTAX05 + WK_C.CSUKUM[iCol];
        //									break;
        //							}
        //						} else {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "H06":
        //									WG03.U_BTXH06 = WG03.U_BTXH06 + YGUCSU;
        //									break;
        //								case "H07":
        //									WG03.U_BTXH07 = WG03.U_BTXH07 + YGUCSU;
        //									break;
        //								case "H08":
        //									WG03.U_BTXH08 = WG03.U_BTXH08 + YGUCSU;
        //									break;
        //								case "H09":
        //									WG03.U_BTXH09 = WG03.U_BTXH09 + YGUCSU;
        //									break;
        //								case "H10":
        //									WG03.U_BTXH10 = WG03.U_BTXH10 + YGUCSU;
        //									break;
        //								default:
        //									BTAX05 = BTAX05 + YGUCSU;
        //									break;
        //							}
        //						}
        //						break;
        //					case "7":
        //						/// 보육수당
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							CHLCSU = CHLCSU + 0;
        //						} else {
        //							CHLCSU = CHLCSU + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & CHLCSU > WK_C.CSUKUM[iCol]) {
        //							WG03.U_BTXQ01 = WG03.U_BTXQ01 + WK_C.CSUKUM[iCol];
        //						} else {
        //							WG03.U_BTXQ01 = WG03.U_BTXQ01 + CHLCSU;
        //						}
        //						break;
        //					case "8":
        //						/// 비과세-기타제출
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							GITBTX = 0;
        //						} else {
        //							GITBTX = Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //							/// 누적한도체크아님 각 수당별한도체크.
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & GITBTX > WK_C.CSUKUM[iCol]) {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "G01":
        //									WG03.U_BTXG01 = WG03.U_BTXG01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H01":
        //									WG03.U_BTXH01 = WG03.U_BTXH01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H05":
        //									WG03.U_BTXH05 = WG03.U_BTXH05 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H11":
        //									WG03.U_BTXH11 = WG03.U_BTXH11 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H12":
        //									WG03.U_BTXH12 = WG03.U_BTXH12 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "H13":
        //									WG03.U_BTXH13 = WG03.U_BTXH13 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "I01":
        //									WG03.U_BTXI01 = WG03.U_BTXI01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "K01":
        //									WG03.U_BTXK01 = WG03.U_BTXK01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "R10":
        //									WG03.U_BTXR10 = WG03.U_BTXR10 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "S01":
        //									WG03.U_BTXS01 = WG03.U_BTXS01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "T01":
        //									WG03.U_BTXT01 = WG03.U_BTXT01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "X01":
        //									WG03.U_BTXX01 = WG03.U_BTXX01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y01":
        //									WG03.U_BTXY01 = WG03.U_BTXY01 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y02":
        //									WG03.U_BTXY02 = WG03.U_BTXY02 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y03":
        //									WG03.U_BTXY03 = WG03.U_BTXY03 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y20":
        //									WG03.U_BTXY20 = WG03.U_BTXY20 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y21":
        //									WG03.U_BTXY21 = WG03.U_BTXY21 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Y22":
        //									WG03.U_BTXY22 = WG03.U_BTXY22 + WK_C.CSUKUM[iCol];
        //									break;
        //								case "Z01":
        //									WG03.U_BTXZ01 = WG03.U_BTXZ01 + WK_C.CSUKUM[iCol];
        //									break;
        //								default:
        //									BTAX04 = BTAX04 + WK_C.CSUKUM[iCol];
        //									break;
        //							}
        //						} else {
        //							switch (Strings.UCase(WK_C.BTXCOD[iCol])) {
        //								case "G01":
        //									WG03.U_BTXG01 = WG03.U_BTXG01 + GITBTX;
        //									break;
        //								case "H01":
        //									WG03.U_BTXH01 = WG03.U_BTXH01 + GITBTX;
        //									break;
        //								case "H05":
        //									WG03.U_BTXH05 = WG03.U_BTXH05 + GITBTX;
        //									break;
        //								case "H11":
        //									WG03.U_BTXH11 = WG03.U_BTXH11 + GITBTX;
        //									break;
        //								case "H12":
        //									WG03.U_BTXH12 = WG03.U_BTXH12 + GITBTX;
        //									break;
        //								case "H13":
        //									WG03.U_BTXH13 = WG03.U_BTXH13 + GITBTX;
        //									break;
        //								case "I01":
        //									WG03.U_BTXI01 = WG03.U_BTXI01 + GITBTX;
        //									break;
        //								case "K01":
        //									WG03.U_BTXK01 = WG03.U_BTXK01 + GITBTX;
        //									break;
        //								case "R10":
        //									WG03.U_BTXR10 = WG03.U_BTXR10 + GITBTX;
        //									break;
        //								case "S01":
        //									WG03.U_BTXS01 = WG03.U_BTXS01 + GITBTX;
        //									break;
        //								case "T01":
        //									WG03.U_BTXT01 = WG03.U_BTXT01 + GITBTX;
        //									break;
        //								case "X01":
        //									WG03.U_BTXX01 = WG03.U_BTXX01 + GITBTX;
        //									break;
        //								case "Y01":
        //									WG03.U_BTXY01 = WG03.U_BTXY01 + GITBTX;
        //									break;
        //								case "Y02":
        //									WG03.U_BTXY02 = WG03.U_BTXY02 + GITBTX;
        //									break;
        //								case "Y03":
        //									WG03.U_BTXY03 = WG03.U_BTXY03 + GITBTX;
        //									break;
        //								case "Y20":
        //									WG03.U_BTXY20 = WG03.U_BTXY20 + GITBTX;
        //									break;
        //								case "Y21":
        //									WG03.U_BTXY21 = WG03.U_BTXY21 + GITBTX;
        //									break;
        //								case "Y22":
        //									WG03.U_BTXY22 = WG03.U_BTXY22 + GITBTX;
        //									break;
        //								case "Z01":
        //									WG03.U_BTXZ01 = WG03.U_BTXZ01 + GITBTX;
        //									break;
        //								default:
        //									BTAX04 = BTAX04 + GITBTX;
        //									break;
        //							}
        //						}
        //						break;
        //					default:
        //						/// 비과세-기타-미제출
        //						if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y") {
        //							GITBTX_N = 0;
        //						} else {
        //							GITBTX_N = Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //							/// 누적한도체크아님 각 수당별한도체크.
        //						}
        //						if (WK_C.CSUKUM[iCol] > 0 & GITBTX_N > WK_C.CSUKUM[iCol]) {
        //							BTAX07 = BTAX07 + WK_C.CSUKUM[iCol];
        //						} else {
        //							BTAX07 = BTAX07 + GITBTX_N;
        //						}
        //						break;
        //				}

        //				//// 2010년부터 폐지되는 비과세
        //				if (oYM >= "201001") {
        //					WG03.U_BTXX01 = 0;
        //					WG03.U_BTXY01 = 0;
        //					WG03.U_BTXY02 = 0;
        //					WG03.U_BTXY20 = 0;
        //				}

        //				/// 상여만일경우 U_BONUSS=상여금 U_TOTPAY=상여금+기타소득 아닐경우 상여금만
        //				if (oJOBTYP == "2" & WK_C.BNSUSE[iCol] == "Y" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "N") {
        //					/// 상여를 적용금액이나 변동금액을 적용하였을경우는 기본급또는 상여금(A04)에 나머지상여금액이 0
        //					if (WK_BNSCHK == true) {
        //						WG03.U_CSUAMT[iCol] = 0;
        //					} else {
        //						/// 상여 상여금에 포함 수당은 총지급액에 합산하지않아야함.
        //						if (WK_APPBNS > 0 & WK_C.CSUCOD[iCol] == "A01") {
        //							WG03.U_CSUAMT[iCol] = WG03.U_BONUSS;
        //							WK_APPBNS = 0;
        //							WK_BNSCHK = true;
        //						/// 기본급이 상여포함구분이 아닐경우 상여금란에 표시
        //						} else if (WK_APPBNS > 0 & WK_C.CSUCOD[iCol] == "A04") {
        //							WG03.U_CSUAMT[iCol] = WG03.U_BONUSS;
        //							WK_APPBNS = 0;
        //							WK_BNSCHK = true;
        //						}
        //					}
        //				} else {
        //					WG03.U_TOTPAY = WG03.U_TOTPAY + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //					/// 월정급여
        //					if (WK_C.MPYGBN[iCol] == "Y") {
        //						MONPAY = MONPAY + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //					}
        //					/// 고용보험대상금액
        //					if (WK_C.GBHGBN[iCol] == "Y") {
        //						TOTGBH = TOTGBH + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[iCol]));
        //					}
        //				}
        //			}
        //			/// 상여일경우 상여금금액 더 적용
        //			if (oJOBTYP == "2") {
        //				WG03.U_TOTPAY = WG03.U_TOTPAY + WG03.U_BONUSS;
        //				MONPAY = MONPAY + WG03.U_BONUSS;
        //				TOTGBH = TOTGBH + WG03.U_BONUSS;
        //			}
        //			if (oJOBTYP != "1" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "Y") {
        //				switch (oJOBTYP) {
        //					case "1":
        //						break;
        //					case "2":
        //						WG03.U_BONUSS = WG03.U_TOTPAY;
        //						WG03.U_APPRAT = Convert.ToDouble("100");
        //						WG03.U_BNSRAT = Convert.ToDouble("100");
        //						break;
        //					case "3":
        //						WG03.U_APPRAT = Convert.ToDouble("100");
        //						WG03.U_BNSRAT = Convert.ToDouble("100");
        //						break;
        //				}
        //			}

        //			WG03.U_TOTPAY = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_TOTPAY, "##########0"));
        //			if (WG03.U_TOTPAY == 0)
        //				return;
        //			// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			/// 공제내역 계산/
        //			// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			/// 비과세산출 /*******************************************************************************/
        //			/// 비과세누적액 산출
        //			WK_YBTXAM[1] = 0;
        //			WK_YBTXAM[2] = 0;
        //			WK_YBTXAM[3] = 0;
        //			WK_YBTXAM[4] = 0;
        //			WK_YBTXAM[5] = 0;
        //			WK_YBTXAM[6] = 0;
        //			WK_YBTXAM[7] = 0;
        //			sQry = "EXEC PH_PY111_BTX '" + oYM + "','" + oJOBTYP + "', '" + oJOBGBN + "', '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount > 0) {
        //				WK_YBTXAM[1] = Conversion.Val(oRecordSet.Fields.Item(0).Value);
        //				/// 연누적액(생산직비과세)
        //				WK_YBTXAM[2] = Conversion.Val(oRecordSet.Fields.Item(1).Value);
        //				/// 월누적(식대,차량보조)
        //				WK_YBTXAM[3] = Conversion.Val(oRecordSet.Fields.Item(2).Value);
        //				/// 월누적(국외비과세)
        //				WK_YBTXAM[4] = Conversion.Val(oRecordSet.Fields.Item(3).Value);
        //				/// 월누적(기타비과세(제출)
        //				WK_YBTXAM[5] = Conversion.Val(oRecordSet.Fields.Item(4).Value);
        //				/// 월누적(연구비과세)
        //				WK_YBTXAM[6] = Conversion.Val(oRecordSet.Fields.Item(5).Value);
        //				/// 월누적(출산보육비과세)
        //				WK_YBTXAM[7] = Conversion.Val(oRecordSet.Fields.Item(6).Value);
        //				/// 월누적-기타비과세(미제출)
        //			}

        //			/// 비과세3-국외비과세대상
        //			/// 비과세대상자만
        //			if (sRecordset.Fields.Item("U_FRGSEL").Value == "Y" & oJOBTYP != "2") {
        //				switch (Strings.Trim(TB1_BT3COD)) {
        //					case "M02":
        //						/// 월한도액(백오십만원)
        //						if (WG03.U_TOTPAY < WG01.TB1AMT[7]) {
        //							WG03.U_BTXM02 = WG03.U_TOTPAY;
        //						} else {
        //							WG03.U_BTXM02 = WG01.TB1AMT[7];
        //						}
        //						if ((WG03.U_BTXM02 + WK_YBTXAM[3]) > WG01.TB1AMT[7])
        //							WG03.U_BTXM02 = WG01.TB1AMT[7] - (WG03.U_BTXM02 + WK_YBTXAM[3]);
        //						if (WG03.U_BTXM02 < 0)
        //							WG03.U_BTXM02 = 0;
        //						break;
        //					case "M03":
        //						/// 한도제한제외
        //						if (WG03.U_TOTPAY < WG03.U_BTXM03) {
        //							WG03.U_BTXM03 = WG03.U_TOTPAY;
        //						}
        //						if (WG03.U_BTXM03 < 0)
        //							WG03.U_BTXM03 = 0;
        //						break;
        //					default:
        //						/// M01
        //						/// 월한도액(백만원)
        //						if (WG03.U_TOTPAY < WG01.TB1AMT[4]) {
        //							WG03.U_BTXM01 = WG03.U_TOTPAY;
        //						} else {
        //							WG03.U_BTXM01 = WG01.TB1AMT[4];
        //						}
        //						if ((WG03.U_BTXM01 + WK_YBTXAM[3]) > WG01.TB1AMT[4])
        //							WG03.U_BTXM01 = WG01.TB1AMT[4] - (WG03.U_BTXM01 + WK_YBTXAM[3]);
        //						if (WG03.U_BTXM01 < 0)
        //							WG03.U_BTXM01 = 0;
        //						break;
        //				}
        //			}
        //			GWASEE = WG03.U_TOTPAY - WG03.U_BTXM01 - WG03.U_BTXM02 - WG03.U_BTXM03;

        //			/// 비과세1-생산직 비과세
        //			/// 비과세대상자만
        //			if (sRecordset.Fields.Item("U_BX1SEL").Value == "Y") {
        //				if ((MONPAY - BT1AMT) <= 1000000) {
        //					WG03.U_BTXO01 = BT1AMT;
        //					//비과세 해당자
        //				} else {
        //					WG03.U_BTXO01 = 0;
        //				}
        //				if (BT1KUM != 0 & WG03.U_BTXO01 > BT1KUM)
        //					WG03.U_BTXO01 = BT1KUM;
        //				/// 한도액구함
        //				/// 연간 누적 비과세1 체크(상여계산시는 제외)
        //				if (oJOBTYP != "2") {
        //					/// (비과세1+연누적비과세)> 연240만원 THEN 비과세1 = 연240만원-(연누적비과세+비과세1)
        //					if ((WG03.U_BTXO01 + WK_YBTXAM[1]) > WG01.TB1KUM[1]) {
        //						WG03.U_BTXO01 = WG01.TB1KUM[1] - WK_YBTXAM[1];
        //					}
        //					if (WG03.U_BTXO01 < 0)
        //						BTAX01 = 0;
        //				}
        //			}

        //			/// 2009년이전자료이면 이전비과세란으로
        //			if (oYM < "200901") {
        //				/// 생산직비과세
        //				BTAX01 = WG03.U_BTXO01;
        //				/// 국외비과세
        //				BTAX03 = WG03.U_BTXM01 + WG03.U_BTXM02 + WG03.U_BTXM03;
        //				/// 연구개발
        //				BTAX05 = WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09 + WG03.U_BTXH10;
        //				/// 출산보육
        //				BTAX06 = WG03.U_BTXQ01;
        //				/// 기타지급조서제출분
        //				if (oYM < "201001") {
        //					BTAX04 = WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXX01 + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY20 + WG03.U_BTXZ01;
        //				} else {
        //					BTAX04 = WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY21 + WG03.U_BTXZ01;
        //				}
        //				WG03.U_BTXG01 = 0;
        //				WG03.U_BTXH01 = 0;
        //				WG03.U_BTXH05 = 0;
        //				WG03.U_BTXH06 = 0;
        //				WG03.U_BTXH07 = 0;
        //				WG03.U_BTXH08 = 0;
        //				WG03.U_BTXH09 = 0;
        //				WG03.U_BTXH10 = 0;
        //				WG03.U_BTXH11 = 0;
        //				WG03.U_BTXH12 = 0;
        //				WG03.U_BTXH13 = 0;
        //				WG03.U_BTXI01 = 0;
        //				WG03.U_BTXK01 = 0;
        //				WG03.U_BTXM01 = 0;
        //				WG03.U_BTXM02 = 0;
        //				WG03.U_BTXM03 = 0;
        //				WG03.U_BTXO01 = 0;
        //				WG03.U_BTXQ01 = 0;
        //				WG03.U_BTXR10 = 0;
        //				WG03.U_BTXS01 = 0;
        //				WG03.U_BTXT01 = 0;
        //				WG03.U_BTXX01 = 0;
        //				WG03.U_BTXY01 = 0;
        //				WG03.U_BTXY02 = 0;
        //				WG03.U_BTXY03 = 0;
        //				WG03.U_BTXY20 = 0;
        //				WG03.U_BTXY21 = 0;
        //				WG03.U_BTXY22 = 0;
        //				WG03.U_BTXZ01 = 0;
        //			}

        //			/// 1.2) 과세대상급여 ( 총급여액-비과세소득 )
        //			if (GWASEE < BTAX01)
        //				BTAX01 = GWASEE;
        //			GWASEE = GWASEE - BTAX01;
        //			/// 비과세2-식대차량보조
        //			if (GWASEE < BTAX02)
        //				BTAX02 = GWASEE;
        //			GWASEE = GWASEE - BTAX02;
        //			/// 비과세4
        //			if (GWASEE < BTAX04)
        //				BTAX04 = GWASEE;
        //			GWASEE = GWASEE - BTAX04;
        //			/// 비과세5
        //			if (GWASEE < BTAX05)
        //				BTAX05 = GWASEE;
        //			GWASEE = GWASEE - BTAX05;
        //			/// 비과세6
        //			if (GWASEE < BTAX06)
        //				BTAX06 = GWASEE;
        //			GWASEE = GWASEE - BTAX06;
        //			/// 비과세7
        //			if (GWASEE < BTAX07)
        //				BTAX07 = GWASEE;
        //			GWASEE = GWASEE - BTAX07;
        //			/// 비과세 (G01~Z01)
        //			if (GWASEE < WG03.U_BTXG01)
        //				WG03.U_BTXG01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXG01;
        //			if (GWASEE < WG03.U_BTXH01)
        //				WG03.U_BTXH01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH01;
        //			if (GWASEE < WG03.U_BTXH05)
        //				WG03.U_BTXH05 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH05;
        //			if (GWASEE < WG03.U_BTXH06)
        //				WG03.U_BTXH06 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH06;
        //			if (GWASEE < WG03.U_BTXH07)
        //				WG03.U_BTXH07 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH07;
        //			if (GWASEE < WG03.U_BTXH08)
        //				WG03.U_BTXH08 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH08;
        //			if (GWASEE < WG03.U_BTXH09)
        //				WG03.U_BTXH09 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH09;
        //			if (GWASEE < WG03.U_BTXH10)
        //				WG03.U_BTXH10 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH10;
        //			if (GWASEE < WG03.U_BTXH11)
        //				WG03.U_BTXH11 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH11;
        //			if (GWASEE < WG03.U_BTXH12)
        //				WG03.U_BTXH12 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH12;
        //			if (GWASEE < WG03.U_BTXH13)
        //				WG03.U_BTXH13 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXH13;
        //			if (GWASEE < WG03.U_BTXI01)
        //				WG03.U_BTXI01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXI01;
        //			if (GWASEE < WG03.U_BTXK01)
        //				WG03.U_BTXK01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXK01;
        //			if (GWASEE < WG03.U_BTXO01)
        //				WG03.U_BTXO01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXO01;
        //			if (GWASEE < WG03.U_BTXQ01)
        //				WG03.U_BTXQ01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXQ01;
        //			if (GWASEE < WG03.U_BTXR10)
        //				WG03.U_BTXR10 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXR10;
        //			if (GWASEE < WG03.U_BTXS01)
        //				WG03.U_BTXS01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXS01;
        //			if (GWASEE < WG03.U_BTXT01)
        //				WG03.U_BTXT01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXT01;
        //			if (GWASEE < WG03.U_BTXX01)
        //				WG03.U_BTXX01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXX01;
        //			if (GWASEE < WG03.U_BTXY01)
        //				WG03.U_BTXY01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY01;
        //			if (GWASEE < WG03.U_BTXY02)
        //				WG03.U_BTXY02 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY02;
        //			if (GWASEE < WG03.U_BTXY03)
        //				WG03.U_BTXY03 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY03;
        //			if (GWASEE < WG03.U_BTXY20)
        //				WG03.U_BTXY20 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY20;
        //			if (GWASEE < WG03.U_BTXY21)
        //				WG03.U_BTXY21 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY21;
        //			if (GWASEE < WG03.U_BTXY22)
        //				WG03.U_BTXY22 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXY22;
        //			if (GWASEE < WG03.U_BTXZ01)
        //				WG03.U_BTXZ01 = GWASEE;
        //			GWASEE = GWASEE - WG03.U_BTXZ01;

        //			/// 1.2) 과세대상급여 ( 총급여액-비과세소득 )
        //			WG03.U_GWASEE = GWASEE;
        //			WG03.U_BTAX01 = BTAX01;
        //			WG03.U_BTAX02 = BTAX02;
        //			WG03.U_BTAX03 = BTAX03;
        //			WG03.U_BTAX04 = BTAX04;
        //			WG03.U_BTAX05 = BTAX05;
        //			WG03.U_BTAX06 = BTAX06;
        //			WG03.U_BTAX07 = BTAX07;

        //			/// 비과세합계
        //			if (oYM < "201001") {
        //				WG03.U_BTXTOT = WG03.U_BTAX01 + WG03.U_BTAX02 + WG03.U_BTAX03 + WG03.U_BTAX04 + WG03.U_BTAX05 + WG03.U_BTAX06 + WG03.U_BTAX07;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXH10 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXM01;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXM02 + WG03.U_BTXM03 + WG03.U_BTXO01 + WG03.U_BTXQ01 + WG03.U_BTXS01 + WG03.U_BTXT01 + WG03.U_BTXX01;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY20 + WG03.U_BTXZ01;
        //			} else {
        //				WG03.U_BTXTOT = WG03.U_BTAX01 + WG03.U_BTAX02 + WG03.U_BTAX03 + WG03.U_BTAX04 + WG03.U_BTAX05 + WG03.U_BTAX06 + WG03.U_BTAX07;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXG01 + WG03.U_BTXH01 + WG03.U_BTXH05 + WG03.U_BTXH06 + WG03.U_BTXH07 + WG03.U_BTXH08 + WG03.U_BTXH09;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXH10 + WG03.U_BTXH11 + WG03.U_BTXH12 + WG03.U_BTXH13 + WG03.U_BTXI01 + WG03.U_BTXK01 + WG03.U_BTXM01;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXM02 + WG03.U_BTXM03 + WG03.U_BTXO01 + WG03.U_BTXQ01 + WG03.U_BTXR10 + WG03.U_BTXS01 + WG03.U_BTXT01;
        //				WG03.U_BTXTOT = WG03.U_BTXTOT + WG03.U_BTXY01 + WG03.U_BTXY02 + WG03.U_BTXY03 + WG03.U_BTXY21 + WG03.U_BTXY22 + WG03.U_BTXZ01;
        //			}

        //			/// 세액계산 /*******************************************************************************/
        //			/// 세액계산안함으로 하면 소득세,주민세 계산안함.
        //			if (oDS_PH_PY111A.GetValue("U_TAXCHK", 0) == "N") {
        //				/// 총지급액, 과세대상금액, 공제인원수
        //				GABGUN = 0;
        //				JUMINN = 0;

        //			} else {
        //				/// 1.3) 세액계산
        //				if (sRecordset.Fields.Item("U_TAXSEL").Value == "Y") {
        //					INCOME = GWASEE;
        //					//- IIf(sRecordSet.Fields("U_BX1SEL").Value = "Y", BNSCSU, 0)
        //					/// (2010.01.19) 급상여이면 세액기간에+1하던걸 세액기간이 없을때만 +1로변경함
        //					//If oJOBTYP = "3" Then WG03.U_TAXTRM = WG03.U_TAXTRM + 1
        //					if (oJOBTYP == "3" & WG03.U_TAXTRM == 0)
        //						WG03.U_TAXTRM = WG03.U_TAXTRM + 1;
        //					/// 상여처리
        //					if (oJOBTYP != "1") {
        //						/// (2010.01.19) 변동자료만 적용버튼추가되면서 세액기간없을경우 1로 셋팅하도록함.
        //						if (WG03.U_TAXTRM == 0)
        //							WG03.U_TAXTRM = 1;
        //						//// (2011.10.05) 상여금 급여기본등록의 평균임금참조로 설정시
        //						if (Strings.Trim(PAY_007) == "2") {
        //							WG03.U_AVRPAY = WK_AVRAMT * WG03.U_TAXTRM;
        //						}
        //						//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						INCOME = MDC_SetMod.IInt(ref (GWASEE + WG03.U_AVRPAY) / WG03.U_TAXTRM, ref 1);
        //					}

        //					/// 총지급액, 과세대상금액, 공제인원수
        //					if (Strings.Trim(PAY_001) == "4") {
        //						MDC_SetMod.Get_GabGunSe_Table(ref GABGUN, ref JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001);
        //					} else {
        //						MDC_SetMod.Get_GabGunSe(ref GABGUN, ref JUMINN, INCOME, WG03.U_TAXCNT, WG03.U_CHLCNT, oYM, INCOME, PAY_001);
        //					}
        //				} else {
        //					GABGUN = 0;
        //					JUMINN = 0;
        //				}
        //			}
        //			///****************************************************************************************************************************/
        //			/// 나이체크
        //			/// 고용보험(만18세이상~만65미만가입가능) 64세이상월부터 고용보험제외(200805월기준 1944.5.25일생일경우 5월분 고용보험까지납부,6월분부터 공제제외)
        //			/// 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
        //			///****************************************************************************************************************************/
        //			KUKSTR = "";
        //			KUKEND = "";
        //			GBHSTR = "";
        //			GBHEND = "";
        //			if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_PERNBR))) {
        //				switch (Strings.Mid(Strings.Trim(WG03.U_PERNBR), 7, 1)) {
        //					case "1":
        //					case "2":
        //					case "5":
        //					case "6":
        //						KUKSTR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format("19" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01", "0000-00-00");
        //						KUKEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, 60, Convert.ToDateTime(KUKSTR)), "yyyymm");
        //						GBHSTR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format("19" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01", "0000-00-00");
        //						GBHEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, 64, Convert.ToDateTime(GBHSTR)), "yyyymm");
        //						break;
        //					case "9":
        //					case "0":
        //						KUKSTR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format("18" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01", "0000-00-00");
        //						KUKEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, 60, Convert.ToDateTime(KUKSTR)), "yyyymm");
        //						GBHSTR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format("18" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01", "0000-00-00");
        //						GBHEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, 64, Convert.ToDateTime(GBHSTR)), "yyyymm");
        //						break;
        //					case "3":
        //					case "4":
        //					case "7":
        //					case "8":
        //						KUKSTR = "20" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01";
        //						KUKEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, 60, Convert.ToDateTime(KUKSTR)), "yyyymm");
        //						GBHSTR = "20" + Strings.Mid(Strings.Trim(WG03.U_PERNBR), 1, 4) + "01";
        //						GBHEND = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, 64, Convert.ToDateTime(GBHSTR)), "yyyymm");
        //						break;
        //					default:
        //						KUKSTR = "";
        //						KUKEND = "";
        //						GBHSTR = "";
        //						GBHEND = "";
        //						break;
        //				}
        //			}

        //			/// 1.4) 고용보험
        //			//// 고용보험 계산안함으로 하면 제외(단, 체크되어 있으면 무조건 계산)
        //			//2012.07.11 대우(동우공영:DWGY)의 요청으로 고용보험료의 산출방식이 "고용보험 보수월액기준"이면서 중도입사자일때 일할 계산 처리 추가
        //			//고용보험 보수월액의 일할 계산(중도입사자) = 고용보험 보수월액 / 월총일수 * 근무일수
        //			//2012.07.13 대우(동우공영:DWGY)의 요청으로 고용보험료의 산출방식이 "고용보험 보수월액기준"이면서 중도입사자일때 30일로 계산 수정요청
        //			//고용보험 보수월액의 일할 계산(중도입사자) = 고용보험 보수월액 / 30 * 근무일수
        //			if (oDS_PH_PY111A.GetValue("U_GBHCHK", 0) == "N") {
        //				GBHAMT = 0;
        //			} else {
        //				//// 고용보험율 등록 데이터 확인
        //				sQry = "SELECT TOP 1 U_EMPRAT1, ISNULL(U_EXECHK,'1') AS U_EXECHK FROM [@ZPY118H] ";
        //				sQry = sQry + " WHERE CODE <= '" + oYM + "' ORDER BY CODE DESC";
        //				oRecordSet.DoQuery(sQry);

        //				if (oRecordSet.RecordCount > 0) {
        //					//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					GBHRAT = oRecordSet.Fields.Item("U_EMPRAT1").Value;
        //					//// 보험율
        //					switch (oRecordSet.Fields.Item("U_EXECHK").Value) {
        //						//// 보험료 계산대상 소득금액
        //						case "2":
        //							WK_GBHAMT = TOTGBH;
        //							break;
        //						case "3":
        //							WK_GBHAMT = WG03.U_GWASEE;
        //							break;
        //						//2012.07.11 대우(동우공영:DWGY) 요청
        //						//Case Else:  WK_GBHAMT = WG03.U_GBHAMT
        //						default:
        //							WK_GBHAMT = WG03.U_GBHAMT;
        //							break;
        //					}
        //				//// 고용보험율 등록 데이터가 없으면 내부 기본값으로 계산
        //				} else {
        //					if (oYM < "201104") {
        //						GBHRAT = 0.45;
        //						if (oYM < "201101") {
        //							WK_GBHAMT = TOTGBH;
        //							//// 2011년 이전은 보수총액
        //						} else {
        //							WK_GBHAMT = WG03.U_GBHAMT;
        //							//// 2011년 이후는 보수월액
        //						}
        //					} else {
        //						GBHRAT = 0.55;
        //						WK_GBHAMT = WG03.U_GBHAMT;
        //					}
        //				}

        //				if (sRecordset.Fields.Item("U_GBHSEL").Value == "Y") {
        //					//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					GBHAMT = MDC_SetMod.IInt(ref WK_GBHAMT * (GBHRAT / 100), ref 10);

        //					/// 고용보험제외(~만65미만가입가능) 64세이상월부터 고용보험제외(200805월기준 1943.5.25일생일경우 4월분 고용보험까지납부,5월분부터 공제제외)
        //					if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_PERNBR))) {
        //						if (!string.IsNullOrEmpty(Strings.Trim(GBHEND)) & Strings.Trim(GBHEND) <= oYM) {
        //							GBHAMT = 0;
        //							REMARK1 = REMARK1 + Strings.Space(1) + WG03.U_MSTCOD + WG03.U_MSTNAM;
        //						}
        //					}
        //				}
        //			}

        //			//// 건강보험과 국민연금은 급여+정기, 급상여+정기일때만 자동계산
        //			//// 2010.08.18 최동권 수정. 일성신약은 상여일 때도 건강/장기요양 보험료 계산(과세금액 기준)
        //			if ((oJOBTYP == "1" & oJOBGBN == "1") | (oJOBTYP == "3" & oJOBGBN == "1") | G04_BNSUSE == "Y") {
        //				WK_MEDAMT = 0;
        //				/// 1.5) 의료보험 /'/ 당월 1일이후 입사자 건강보험료 제외
        //				if (Conversion.Val(Convert.ToString(WG03.U_MEDAMT)) != 0) {
        //					sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO, U_NJYYMM, U_NJYRAT, U_NJCRAT, ISNULL(U_EXECHK,1) AS U_EXECHK FROM [@ZPY103H] ";
        //					sQry = sQry + " WHERE CODE <= '" + oYM + "' ORDER BY CODE DESC";
        //					oRecordSet.DoQuery(sQry);
        //					if (oRecordSet.RecordCount > 0) {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJYYMM = oRecordSet.Fields.Item("U_NJYYMM").Value;
        //						/// 장기요양보험적용월
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJYRAT = oRecordSet.Fields.Item("U_NJYRAT").Value;
        //						/// 장기요양보험적용율
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJCRAT = oRecordSet.Fields.Item("U_NJCRAT").Value;
        //						/// 장기요양보험경감보험요율
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						MEDCHK = oRecordSet.Fields.Item("U_EXECHK").Value;
        //						/// 건강보험산출방식구분(1:보수월액기준,2:월과세금액기준,3:지급총액기준)
        //						switch (Strings.Trim(MEDCHK)) {
        //							case "2":
        //								WK_MEDAMT = WG03.U_GWASEE;
        //								/// 과세금액기준
        //								break;
        //							case "3":
        //								WK_MEDAMT = WG03.U_TOTPAY;
        //								/// 총지급액기준
        //								break;
        //							default:
        //								WK_MEDAMT = Conversion.Val(Convert.ToString(WG03.U_MEDAMT));
        //								break;
        //						}

        //						if (Conversion.Val(Convert.ToString(WK_MEDAMT)) < oRecordSet.Fields.Item("U_FROM").Value) {
        //							//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							MEDAMT = oRecordSet.Fields.Item("U_FROM").Value;
        //						} else if (oRecordSet.Fields.Item("U_TO").Value > 0 & Conversion.Val(Convert.ToString(WK_MEDAMT)) > oRecordSet.Fields.Item("U_TO").Value) {
        //							//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							MEDAMT = oRecordSet.Fields.Item("U_TO").Value;
        //						} else {
        //							MEDAMT = Conversion.Val(Convert.ToString(WK_MEDAMT));
        //						}
        //						//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						MEDAMT = MDC_SetMod.IInt(ref MEDAMT * Conversion.Val(oRecordSet.Fields.Item("U_EMPRAT").Value) / 100, ref 10);
        //						/// 해외파견경감대상자일경우
        //						if (Conversion.Val(Convert.ToString(WG03.U_MEDFRG)) != 0) {
        //							MEDAMT = MEDAMT - (MDC_SetMod.IInt(ref MEDAMT * (WG03.U_MEDFRG / 100) + 9, ref 10));
        //							/// 경감보험료(10원미만절상)
        //						}
        //					}
        //					if (oYM == Strings.Mid(Strings.Replace(WG03.U_INPDAT, "-", ""), 1, 6)) {
        //						if (oYM + "01" != Strings.Mid(WG03.U_INPDAT, 1, 8)) {
        //							MEDAMT = 0;
        //							/// 당월 1일이후 입사자 건강보험료 제외
        //						}
        //					}
        //				}
        //				/// 1.5.5) 노인장기요양보험료율
        //				if (MEDAMT != 0) {
        //					sQry = "SELECT TOP 1 U_NJYYMM, U_NJYRAT, U_NJCRAT FROM [@ZPY103H] ";
        //					sQry = sQry + " WHERE U_NJYYMM <= '" + oYM + "' ORDER BY U_NJYYMM DESC";
        //					oRecordSet.DoQuery(sQry);
        //					if (oRecordSet.RecordCount > 0) {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJYYMM = oRecordSet.Fields.Item("U_NJYYMM").Value;
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJYRAT = oRecordSet.Fields.Item("U_NJYRAT").Value;
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						U_NJCRAT = oRecordSet.Fields.Item("U_NJCRAT").Value;
        //					} else {
        //						U_NJYYMM = "";
        //						U_NJYRAT = 0;
        //						U_NJCRAT = 0;
        //					}
        //				}
        //			}

        //			if ((oJOBTYP == "1" & oJOBGBN == "1") | (oJOBTYP == "3" & oJOBGBN == "1")) {
        //				/// 1.6) 국민연금 /
        //				if (oYM <= "200803") {
        //					sQry = "SELECT TOP 1 ISNULL(T0.U_EMPAMT, 0) FROM [@ZPY102L] T0 WHERE T0.Code <= '" + oYM + "' ";
        //					sQry = sQry + " AND T0.U_CODNBR ='" + sRecordset.Fields.Item("U_KUKGRD").Value + "'";
        //					sQry = sQry + " ORDER BY CODE DESC";
        //					oRecordSet.DoQuery(sQry);
        //					if (oRecordSet.RecordCount > 0) {
        //						KUKAMT = Conversion.Val(oRecordSet.Fields.Item(0).Value);
        //					}
        //				} else {
        //					/// 국민연금보수월액이랑 등급없으면 연금공제 제외
        //					if (string.IsNullOrEmpty(Strings.Trim(sRecordset.Fields.Item("U_KUKGRD").Value)) & Conversion.Val(Convert.ToString(WG03.U_KUKAMT)) == 0) {
        //						KUKAMT = 0;
        //					} else {
        //						sQry = "SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] ";
        //						sQry = sQry + " WHERE CODE <= '" + oYM + "' ORDER BY CODE DESC";
        //						oRecordSet.DoQuery(sQry);
        //						if (oRecordSet.RecordCount > 0) {
        //							if (Conversion.Val(Convert.ToString(WG03.U_KUKAMT)) == 0) {
        //								KUKAMT = 0;
        //							} else if (Conversion.Val(Convert.ToString(WG03.U_KUKAMT)) < oRecordSet.Fields.Item("U_FROM").Value) {
        //								//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								KUKAMT = oRecordSet.Fields.Item("U_FROM").Value;
        //							} else if (oRecordSet.Fields.Item("U_TO").Value > 0 & Conversion.Val(Convert.ToString(WG03.U_KUKAMT)) > oRecordSet.Fields.Item("U_TO").Value) {
        //								//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								KUKAMT = oRecordSet.Fields.Item("U_TO").Value;
        //							} else {
        //								KUKAMT = Conversion.Val(Convert.ToString(WG03.U_KUKAMT));
        //							}
        //							//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							KUKAMT = MDC_SetMod.IInt(ref KUKAMT * Conversion.Val(oRecordSet.Fields.Item("U_EMPRAT").Value) / 100, ref 10);
        //						}
        //					}
        //				}
        //				/// 국민연금(만18세이상~만60미만가입가능) 60세이상월부터 국민연금제외(200805월기준 1948.5.25일생일경우 5월분 국민연금까지납부,6월분부터 공제제외)
        //				/// 예외:급여기본등록의 국민연금 연장에 체크되어있으신분은 나이제한공제제외대상에서 빠짐(국민연금공제)
        //				if (!string.IsNullOrEmpty(Strings.Trim(WG03.U_PERNBR))) {
        //					if (!string.IsNullOrEmpty(Strings.Trim(KUKEND)) & Strings.Trim(KUKEND) < oYM & sRecordset.Fields.Item("U_KUKOVR").Value == "N") {
        //						KUKAMT = 0;
        //						REMARK2 = REMARK2 + Strings.Space(1) + WG03.U_MSTCOD + WG03.U_MSTNAM;
        //					}
        //				}
        //				if (oYM == Strings.Mid(Strings.Replace(WG03.U_INPDAT, "-", ""), 1, 6)) {
        //					if (oYM + "01" != Strings.Mid(WG03.U_INPDAT, 1, 8)) {
        //						KUKAMT = 0;
        //						/// 당월 1일이후 입사자 국민연금 제외
        //					}
        //				}

        //			}

        //			//// 소득세, 주민세, 고용보험은 계산여부를 별도로 선택가능하므로
        //			//// 수당/공제 계산식 반영제외와 무관하게 처리함
        //			for (iCol = 1; iCol <= 18; iCol++) {
        //				switch (WK_G.GONCOD[iCol]) {
        //					case "G01":
        //						/// 소득세
        //						if (oJOBTYP == "1") {
        //							//UPGRADE_WARNING: MDC_SetMod.IInt(GABGUN, 10) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + MDC_SetMod.IInt(ref GABGUN, ref 10);
        //						} else {
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + (MDC_SetMod.IInt(ref (GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, ref 10));
        //						}
        //						if (WG03.U_GONAMT[iCol] < 0)
        //							WG03.U_GONAMT[iCol] = 0;
        //						break;
        //					case "G02":
        //						/// 주민세
        //						if (oJOBTYP == "1") {
        //							//UPGRADE_WARNING: MDC_SetMod.IInt(JUMINN, 10) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + MDC_SetMod.IInt(ref JUMINN, ref 10);
        //						} else {
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + (MDC_SetMod.IInt(ref (MDC_SetMod.IInt(ref (GABGUN * WG03.U_TAXTRM) - WG03.U_NABTAX, ref 10)) * 0.1, ref 10));
        //						}
        //						if (WG03.U_GONAMT[iCol] < 0)
        //							WG03.U_GONAMT[iCol] = 0;
        //						break;
        //					case "G05":
        //						/// 고용보험
        //						WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + GBHAMT;
        //						break;
        //					case "G04":
        //						/// 건강보험
        //						if (G04_BNSUSE == "Y" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "Y") {
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + MEDAMT;
        //						}
        //						break;
        //					case "G10":
        //						/// 장기요양보험
        //						if (G04_BNSUSE == "Y" & oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "Y") {
        //							if (U_NJYYMM <= oYM) {
        //								//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								NJCMED = MDC_SetMod.IInt(ref MEDAMT * (U_NJYRAT / 100), ref 10);
        //								/// (10원미만 절사)
        //								/// 경감대상자일경우
        //								if (Strings.Trim(WG03.U_NJCGBN) == "Y") {
        //									NJCMED = NJCMED - (MDC_SetMod.IInt(ref NJCMED * (U_NJCRAT / 100) + 9, ref 10));
        //									/// 경감보험료(10원미만절상)
        //								}
        //								WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + NJCMED;
        //								/// 장기요양보험료
        //							}
        //						}
        //						break;
        //				}
        //			}

        //			//// 2. 공식가지고 계산하기-고정공제
        //			if (oDS_PH_PY111A.GetValue("U_CHGCHK", 0) == "N") {
        //				for (iCol = 1; iCol <= 18; iCol++) {
        //					switch (WK_G.GONCOD[iCol]) {
        //						case "G03":
        //							/// 국민연금
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + KUKAMT;
        //							break;
        //						case "G04":
        //							/// 건강보험
        //							WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + MEDAMT;
        //							break;
        //						case "G10":
        //							/// 노인장기요양보험
        //							if (U_NJYYMM <= oYM) {
        //								//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								NJCMED = MDC_SetMod.IInt(ref MEDAMT * (U_NJYRAT / 100), ref 10);
        //								/// (10원미만 절사)
        //								/// 경감대상자일경우
        //								if (Strings.Trim(WG03.U_NJCGBN) == "Y") {
        //									NJCMED = NJCMED - (MDC_SetMod.IInt(ref NJCMED * (U_NJCRAT / 100) + 9, ref 10));
        //									/// 경감보험료(10원미만절상)
        //								}
        //								WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + NJCMED;
        //								/// 장기요양보험료
        //							}
        //							break;
        //						default:
        //							/// 상여만일경우 공제체크안된것 제외
        //							if (oJOBTYP == "2" & WK_G.BNSUSE[iCol] == "N") {
        //								WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + 0;
        //							} else {
        //								/// 공식이있으면 공식계산
        //								if (!string.IsNullOrEmpty(Strings.Trim(WK_G.GONCOD[iCol])) & !string.IsNullOrEmpty(Strings.Trim(WK_G.GONSIL[iCol]))) {
        //									/// 2.1. 공식-계산결과값가져오는거면..
        //									SuSilStr = Change_GOSIL(ref WK_G.GONSIL[iCol]);
        //									/// 계산된값일 경우
        //									for (kCol = 1; kCol <= 18; kCol++) {
        //										SuSilStr = Strings.Replace(SuSilStr, "#" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(kCol, "00"), Convert.ToString(WG03.U_GONAMT[kCol]));
        //										/// 계산된 공제값
        //									}
        //									for (kCol = 1; kCol <= 24; kCol++) {
        //										SuSilStr = Strings.Replace(SuSilStr, "#C" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(kCol, "00"), Convert.ToString(WG03.U_CSUAMT[kCol]));
        //										/// 계산된 수당값
        //									}
        //									SuSilStr = Strings.Replace(SuSilStr, "X15", Convert.ToString(WG03.U_TOTPAY));
        //									/// X15:지급총액
        //									/// 2.2. 공식계산하기
        //									Tmp_CSUAMT = Get_ReAmt(ref oYM, ref WG03.U_MSTCOD, ref SuSilStr);
        //									/// 2.3. 정답가져오기
        //									//UPGRADE_WARNING: MDC_SetMod.RInt(Tmp_CSUAMT, WK_G.RODLEN(iCol), WK_G.ROUNDT(iCol)) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									WG03.U_GONAMT[iCol] = WG03.U_GONAMT[iCol] + MDC_SetMod.RInt(ref Tmp_CSUAMT, WK_G.RODLEN[iCol], WK_G.ROUNDT[iCol]);
        //								}
        //							}
        //							break;
        //					}
        //				}
        //			}
        //			/// 급상여변동자료(기간) -변동금액으로 대체/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //			sQry = "SELECT  T0.U_CSUTYP, T0.U_CSUCOD, T0.U_CSUAMT, T0.U_MSTCOD";
        //			sQry = sQry + " FROM [@ZPY312L] T0 INNER JOIN [@ZPY312H] T1 ON T0.DocEntry = T1.DocEntry";
        //			sQry = sQry + " WHERE T1.U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND   T1.U_JOBGBN = '" + oJOBGBN + "'";
        //			sQry = sQry + " AND   T0.U_CSUTYP = '2'";
        //			//공제항목만
        //			sQry = sQry + " AND   T0.U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			sQry = sQry + " AND   " + "'" + oYM + "' BETWEEN T0.U_STRYMM AND T0.U_ENDYMM";
        //			sQry = sQry + " ORDER BY T0.U_MSTCOD, T1.U_YM DESC, T0.U_STRYMM DESC";
        //			oRecordSet.DoQuery(sQry);
        //			while (!(oRecordSet.EoF)) {
        //				for (iCol = 1; iCol <= 18; iCol++) {
        //					if (oRecordSet.Fields.Item("U_CSUCOD").Value == WK_G.GONCOD[iCol]) {
        //						//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						WG03.U_GONAMT[iCol] = oRecordSet.Fields.Item("U_CSUAMT").Value;
        //					}
        //				}
        //				oRecordSet.MoveNext();
        //			}

        //			///ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			///2. 총공제액
        //			///ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			for (iCol = 1; iCol <= 18; iCol++) {
        //				WG03.U_TOTGON = WG03.U_TOTGON + Conversion.Val(Convert.ToString(WG03.U_GONAMT[iCol]));
        //			}
        //			///ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			///3. 실지급액
        //			///ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			//UPGRADE_WARNING: MDC_SetMod.IInt() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			WG03.U_SILJIG = MDC_SetMod.IInt(ref WG03.U_TOTPAY - WG03.U_TOTGON, ref 1);
        //			//If WG03.U_SILJIG < 1000 Then WG03.U_SILJIG = 0

        //			//// 실지급액이 마이너스인 사원 확인요청
        //			if (WG03.U_SILJIG < 0) {
        //				REMARK3 = REMARK3 + Strings.Space(1) + WG03.U_MSTCOD + WG03.U_MSTNAM;
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("PH_PY111A_TAX Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //		}
        #endregion

        #region PH_PY111_Save
        //		private string PH_PY111_Save()
        //		{
        //			string functionReturnValue = null;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;
        //			int DocNum = 0;
        //			short iCol = 0;
        //			string S_TYPNAM = null;
        //			string S_GBNNAM = null;
        //			short ErrNum = 0;
        //			int DfSeries = 0;

        //			DfSeries = MDC_GetData.Get_Series_No(ref "UTD");
        //			/// 2009.01월에 추가

        //			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			S_TYPNAM = oForm.Items.Item("JOBTYP").Specific.Selected.Description;
        //			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			S_GBNNAM = oForm.Items.Item("JOBGBN").Specific.Selected.Description;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			/// 1) Search
        //			sQry = "SELECT ISNULL(U_ENDCHK, 'N') AS ENDCHK FROM [@PH_PY111A]";
        //			sQry = sQry + " WHERE U_JOBGBN =  '" + oJOBGBN + "'";
        //			sQry = sQry + " AND U_JOBTYP = '" + oJOBTYP + "'";
        //			sQry = sQry + " AND U_YM = '" + oYM + "'";
        //			sQry = sQry + " AND U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				/// 2) Insert
        //				// 2.1)Autokey
        //				//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				DocNum = MDC_SetMod.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY111'", ref "");

        //				if (DocNum <= 0)
        //					DocNum = 1;
        //				// 2.2)삽입
        //				sQry = "INSERT INTO [@PH_PY111A] (DocEntry, DocNum, Period, Instance,  Series, Handwrtten, Canceled, Object, UserSign, Transfered, Status, CreateDate, CreateTime, DataSource,";
        //				sQry = sQry + " U_YM, U_JOBTYP, U_JOBGBN, U_MSTCOD, U_RETCHK,U_ENDCHK, U_CSUCOD, U_GONCOD,";
        //				sQry = sQry + " U_JIGBIL, U_MSTNAM, U_EmpID, U_CLTCOD, U_MSTBRK,  U_MSTDPT, U_MSTSTP, U_CLTNAM, U_BRKNAM, U_DPTNAM, U_STPNAM,";
        //				sQry = sQry + "  U_PAYTYP, U_JOBTRG, U_JIGCOD, U_HOBONG, U_STDAMT, U_INPDAT, U_OUTDAT, U_TAXCNT, U_BUYN20, U_DAYAMT, U_BASAMT,";
        //				sQry = sQry + "  U_CSUD01, U_CSUD02, U_CSUD03, U_CSUD04, U_CSUD05, U_CSUD06, U_CSUD07, U_CSUD08,";
        //				sQry = sQry + "  U_CSUD09, U_CSUD10, U_CSUD11, U_CSUD12, U_CSUD13, U_CSUD14, U_CSUD15, U_CSUD16,";
        //				sQry = sQry + "  U_CSUD17, U_CSUD18, U_CSUD19, U_CSUD20, U_CSUD21, U_CSUD22, U_CSUD23, U_CSUD24,";
        //				sQry = sQry + "  U_GWASEE, U_BTAX01, U_BTAX02, U_BTAX03, U_BTAX04, U_BTAX05, U_BTAX06, U_BTAX07,";
        //				sQry = sQry + "  U_BTXG01, U_BTXH01, U_BTXH05, U_BTXH06, U_BTXH07, U_BTXH08, U_BTXH09, U_BTXH10,";
        //				sQry = sQry + "  U_BTXH11, U_BTXH12, U_BTXH13, U_BTXI01, U_BTXK01, U_BTXM01, U_BTXM02, U_BTXM03,";
        //				sQry = sQry + "  U_BTXO01, U_BTXQ01, U_BTXS01, U_BTXT01, U_BTXX01, U_BTXY01, U_BTXY02, U_BTXY03,";
        //				sQry = sQry + "  U_BTXY20, U_BTXZ01, U_BTXTOT,";
        //				sQry = sQry + "  U_TOTPAY, U_TOTGON, U_SILJIG,";
        //				sQry = sQry + "  U_GONG01, U_GONG02, U_GONG03, U_GONG04, U_GONG05, U_GONG06, U_GONG07, U_GONG08,";
        //				sQry = sQry + "  U_GONG09, U_GONG10, U_GONG11, U_GONG12, U_GONG13, U_GONG14, U_GONG15, U_GONG16, U_GONG17, U_GONG18,   ";
        //				sQry = sQry + "  U_AVRPAY, U_NABTAX, U_BNSRAT, U_APPRAT, U_GNSYER,";
        //				sQry = sQry + "  U_GNSMON , U_TAXTRM, U_BONUSS, U_TYPNAM, U_GBNNAM ";
        //				sQry = sQry + " ) VALUES(" + DocNum + ", " + DocNum + "," + "15, 0, '" + DfSeries + "', 'N','N', 'PH_PY111', " + MDC_Globals.oCompany.UserSignature + ", 'N', 'O', '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyy-mm-dd") + "', '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "hhmm") + "', 'I',";
        //				sQry = sQry + "'" + oYM + "', '" + oJOBTYP + "','" + oJOBGBN + "', '" + Strings.Trim(WG03.U_MSTCOD) + "', '" + Strings.Trim(oDS_PH_PY111A.GetValue("U_RETCHK", 0)) + "', 'N', '" + U_CSUCOD + "', '" + U_GONCOD + "',";
        //				sQry = sQry + "'" + Strings.Trim(oJIGBIL) + "', N'" + Strings.Trim(WG03.U_MSTNAM) + "', '" + Strings.Trim(WG03.U_EmpID) + "', '" + Strings.Trim(WG03.U_CLTCOD) + "', '" + Strings.Trim(WG03.U_MSTBRK) + "', '" + Strings.Trim(WG03.U_MSTDPT) + "', '" + Strings.Trim(WG03.U_MSTSTP) + "', N'" + Strings.Trim(WG03.U_CLTNAM) + "', N'" + Strings.Trim(WG03.U_BRKNAM) + "', N'" + Strings.Trim(WG03.U_DPTNAM) + "' , N'" + Strings.Trim(WG03.U_STPNAM) + "',";
        //				sQry = sQry + "'" + Strings.Trim(WG03.U_PAYTYP) + "', '" + Strings.Trim(WG03.U_JOBTRG) + "', '" + Strings.Trim(WG03.U_JIGCOD) + "', '" + Strings.Trim(WG03.U_HOBONG) + "', '" + Strings.Trim(Convert.ToString(WG03.U_STDAMT)) + "', '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00") + "' , ";
        //				sQry = sQry + " CASE WHEN '" + Strings.Trim(WG03.U_OUTDAT) + "'='' THEN NULL ELSE '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_OUTDAT, "0000-00-00") + "' END , '" + Conversion.Val(Convert.ToString(WG03.U_TAXCNT)) + "','" + Conversion.Val(Convert.ToString(WG03.U_CHLCNT)) + "','" + Conversion.Val(Convert.ToString(WG03.U_DAYAMT)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BASAMT)) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[1])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[2])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[3])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[4])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[5])) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[6])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[7])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[8])) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[9])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[10])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[11])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[12])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[13])) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[14])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[15])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[16])) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[17])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[18])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[19])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[20])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[21])) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[22])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[23])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_CSUAMT[24])) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_GWASEE)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX02)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX03)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX04)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX05)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX06)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTAX07)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXG01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH05)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH06)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH07)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXH08)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH09)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH10)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH11)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXH12)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXH13)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXI01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXK01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXM01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXM02)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXM03)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXO01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXQ01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXS01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXT01)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXX01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXY01)) + Conversion.Val(Convert.ToString(WG03.U_BTXR10)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXY02)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXY03)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXY20)) + Conversion.Val(Convert.ToString(WG03.U_BTXY22)) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_BTXZ01)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BTXTOT)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_TOTPAY)) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_TOTGON)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_SILJIG)) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[1])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[2])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[3])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[4])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[5])) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[6])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[7])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[8])) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[9])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[10])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[11])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[12])) + "','" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[13])) + "','" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[14])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[15])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[16])) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[17])) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GONAMT[18])) + "', ";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_AVRPAY)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_NABTAX)) + "' , '" + Conversion.Val(Convert.ToString(WG03.U_BNSRAT)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_APPRAT)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_GNSYER)) + "',";
        //				sQry = sQry + "'" + Conversion.Val(Convert.ToString(WG03.U_GNSMON)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_TAXTRM)) + "', '" + Conversion.Val(Convert.ToString(WG03.U_BONUSS)) + "', N'" + S_TYPNAM + "', N'" + S_GBNNAM + "'";
        //				sQry = sQry + "  )";
        //				oRecordSet.DoQuery(sQry);
        //				// 2.3)Autokey증가
        //				DocNum = DocNum + 1;
        //				sQry = "UPDATE ONNM SET AutoKey = " + DocNum + " WHERE  ObjectCode='PH_PY111'";
        //				oRecordSet.DoQuery(sQry);
        //			} else {
        //				/// 3) Update
        //				if (Strings.Trim(oRecordSet.Fields.Item("ENDCHK").Value) == "Y") {
        //					// 3.1)잠긴 자료
        //					ErrNum = 1;
        //					G_ChkCnt = G_ChkCnt + 1;
        //					goto Error_Message;
        //				}
        //				// 3.2)갱신
        //				sQry = " UPDATE [@PH_PY111A]         ";
        //				sQry = sQry + " SET UpdateDate = '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyy-mm-dd") + "'";
        //				sQry = sQry + " ,   UpdateTime = '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "hhmm") + "'";
        //				sQry = sQry + " ,   U_MSTNAM = N'" + Strings.Trim(WG03.U_MSTNAM) + "'";
        //				sQry = sQry + " ,   UserSign = " + MDC_Globals.oCompany.UserSignature;
        //				sQry = sQry + " ,   U_EmpID = '" + Strings.Trim(WG03.U_EmpID) + "'";
        //				sQry = sQry + " ,   U_ENDCHK = 'N'";
        //				//잠금여부
        //				sQry = sQry + " ,   U_CSUCOD = '" + U_CSUCOD + "'";
        //				sQry = sQry + " ,   U_GONCOD = '" + U_GONCOD + "'";
        //				sQry = sQry + " ,   U_RETCHK = '" + oDS_PH_PY111A.GetValue("U_RETCHK", 0) + "'";
        //				//퇴직금포함
        //				sQry = sQry + " ,   U_JIGBIL = '" + Strings.Trim(oJIGBIL) + "'";
        //				sQry = sQry + " ,   U_CLTCOD = '" + Strings.Trim(WG03.U_CLTCOD) + "'";
        //				sQry = sQry + " ,   U_MSTBRK = '" + Strings.Trim(WG03.U_MSTBRK) + "'";
        //				sQry = sQry + " ,   U_MSTDPT = '" + Strings.Trim(WG03.U_MSTDPT) + "'";
        //				sQry = sQry + " ,   U_MSTSTP = '" + Strings.Trim(WG03.U_MSTSTP) + "'";
        //				sQry = sQry + " ,   U_CLTNAM = N'" + Strings.Trim(WG03.U_CLTNAM) + "'";
        //				sQry = sQry + " ,   U_BRKNAM = N'" + Strings.Trim(WG03.U_BRKNAM) + "'";
        //				sQry = sQry + " ,   U_DPTNAM = N'" + Strings.Trim(WG03.U_DPTNAM) + "'";
        //				sQry = sQry + " ,   U_STPNAM = N'" + Strings.Trim(WG03.U_STPNAM) + "'";
        //				sQry = sQry + " ,   U_PAYTYP = '" + Strings.Trim(WG03.U_PAYTYP) + "'";
        //				sQry = sQry + " ,   U_JOBTRG = '" + Strings.Trim(WG03.U_JOBTRG) + "'";
        //				sQry = sQry + " ,   U_JIGCOD = '" + Strings.Trim(WG03.U_JIGCOD) + "'";
        //				sQry = sQry + " ,   U_HOBONG = '" + Strings.Trim(WG03.U_HOBONG) + "'";
        //				sQry = sQry + " ,   U_STDAMT = " + WG03.U_STDAMT;
        //				sQry = sQry + " ,   U_INPDAT = CASE WHEN '" + Strings.Trim(WG03.U_INPDAT) + "'='' THEN NULL ELSE '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_INPDAT, "0000-00-00") + "' END ";
        //				sQry = sQry + " ,   U_OUTDAT = CASE WHEN '" + Strings.Trim(WG03.U_OUTDAT) + "'='' THEN NULL ELSE '" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(WG03.U_OUTDAT, "0000-00-00") + "' END ";
        //				sQry = sQry + " ,   U_TAXCNT = '" + WG03.U_TAXCNT + "'";
        //				sQry = sQry + " ,   U_BUYN20 = '" + WG03.U_CHLCNT + "'";
        //				sQry = sQry + " ,   U_DAYAMT = '" + WG03.U_DAYAMT + "'";
        //				sQry = sQry + " ,   U_BASAMT = '" + WG03.U_BASAMT + "'";
        //				for (iCol = 1; iCol <= 24; iCol++) {
        //					sQry = sQry + " ,   U_CSUD" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(iCol, "00") + " = " + WG03.U_CSUAMT[iCol];
        //				}
        //				sQry = sQry + " ,   U_GWASEE = '" + WG03.U_GWASEE + "'";
        //				sQry = sQry + " ,   U_BTAX01 = '" + WG03.U_BTAX01 + "'";
        //				sQry = sQry + " ,   U_BTAX02 = '" + WG03.U_BTAX02 + "'";
        //				sQry = sQry + " ,   U_BTAX03 = '" + WG03.U_BTAX03 + "'";
        //				sQry = sQry + " ,   U_BTAX04 = '" + WG03.U_BTAX04 + "'";
        //				sQry = sQry + " ,   U_BTAX05 = '" + WG03.U_BTAX05 + "'";
        //				sQry = sQry + " ,   U_BTAX06 = '" + WG03.U_BTAX06 + "'";
        //				sQry = sQry + " ,   U_BTAX07 = '" + WG03.U_BTAX07 + "'";
        //				sQry = sQry + " ,   U_BTXG01 = '" + WG03.U_BTXG01 + "'";
        //				sQry = sQry + " ,   U_BTXH01 = '" + WG03.U_BTXH01 + "'";
        //				sQry = sQry + " ,   U_BTXH05 = '" + WG03.U_BTXH05 + "'";
        //				sQry = sQry + " ,   U_BTXH06 = '" + WG03.U_BTXH06 + "'";
        //				sQry = sQry + " ,   U_BTXH07 = '" + WG03.U_BTXH07 + "'";
        //				sQry = sQry + " ,   U_BTXH08 = '" + WG03.U_BTXH08 + "'";
        //				sQry = sQry + " ,   U_BTXH09 = '" + WG03.U_BTXH09 + "'";
        //				sQry = sQry + " ,   U_BTXH10 = '" + WG03.U_BTXH10 + "'";
        //				sQry = sQry + " ,   U_BTXH11 = '" + WG03.U_BTXH11 + "'";
        //				sQry = sQry + " ,   U_BTXH12 = '" + WG03.U_BTXH12 + "'";
        //				sQry = sQry + " ,   U_BTXH13 = '" + WG03.U_BTXH13 + "'";
        //				sQry = sQry + " ,   U_BTXI01 = '" + WG03.U_BTXI01 + "'";
        //				sQry = sQry + " ,   U_BTXK01 = '" + WG03.U_BTXK01 + "'";
        //				sQry = sQry + " ,   U_BTXM01 = '" + WG03.U_BTXM01 + "'";
        //				sQry = sQry + " ,   U_BTXM02 = '" + WG03.U_BTXM02 + "'";
        //				sQry = sQry + " ,   U_BTXM03 = '" + WG03.U_BTXM03 + "'";
        //				sQry = sQry + " ,   U_BTXO01 = '" + WG03.U_BTXO01 + "'";
        //				sQry = sQry + " ,   U_BTXQ01 = '" + WG03.U_BTXQ01 + "'";
        //				sQry = sQry + " ,   U_BTXS01 = '" + WG03.U_BTXS01 + "'";
        //				sQry = sQry + " ,   U_BTXT01 = '" + WG03.U_BTXT01 + "'";
        //				sQry = sQry + " ,   U_BTXX01 = '" + WG03.U_BTXX01 + "'";
        //				sQry = sQry + " ,   U_BTXY01 = '" + WG03.U_BTXY01 + WG03.U_BTXR10 + "'";
        //				sQry = sQry + " ,   U_BTXY02 = '" + WG03.U_BTXY02 + "'";
        //				sQry = sQry + " ,   U_BTXY03 = '" + WG03.U_BTXY03 + "'";
        //				sQry = sQry + " ,   U_BTXY20 = '" + WG03.U_BTXY20 + WG03.U_BTXY22 + "'";
        //				sQry = sQry + " ,   U_BTXZ01 = '" + WG03.U_BTXZ01 + "'";
        //				sQry = sQry + " ,   U_BTXTOT = '" + WG03.U_BTXTOT + "'";
        //				sQry = sQry + " ,   U_TOTPAY = '" + WG03.U_TOTPAY + "'";
        //				for (iCol = 1; iCol <= 18; iCol++) {
        //					sQry = sQry + " ,   U_GONG" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(iCol, "00") + " =  '" + WG03.U_GONAMT[iCol] + "'";
        //				}
        //				sQry = sQry + " ,   U_TOTGON = '" + WG03.U_TOTGON + "'";
        //				sQry = sQry + " ,   U_SILJIG = '" + WG03.U_SILJIG + "'";
        //				sQry = sQry + " ,   U_AVRPAY = '" + WG03.U_AVRPAY + "'";
        //				sQry = sQry + " ,   U_NABTAX = '" + WG03.U_NABTAX + "'";
        //				sQry = sQry + " ,   U_BNSRAT = '" + WG03.U_BNSRAT + "'";
        //				sQry = sQry + " ,   U_APPRAT = '" + WG03.U_APPRAT + "'";
        //				sQry = sQry + " ,   U_GNSYER = '" + WG03.U_GNSYER + "'";
        //				sQry = sQry + " ,   U_GNSMON = '" + WG03.U_GNSMON + "'";
        //				sQry = sQry + " ,   U_TAXTRM = '" + WG03.U_TAXTRM + "'";
        //				sQry = sQry + " ,   U_BONUSS = '" + WG03.U_BONUSS + "'";
        //				sQry = sQry + " ,   U_TYPNAM = N'" + S_TYPNAM + "'";
        //				sQry = sQry + " ,   U_GBNNAM = N'" + S_GBNNAM + "'";
        //				sQry = sQry + " WHERE U_YM = '" + oYM + "'";
        //				sQry = sQry + " AND   U_JOBTYP = '" + oJOBTYP + "'";
        //				sQry = sQry + " AND   U_JOBGBN = '" + oJOBGBN + "'";
        //				sQry = sQry + " AND   U_MSTCOD = '" + Strings.Trim(WG03.U_MSTCOD) + "'";
        //				oRecordSet.DoQuery(sQry);

        //			}
        //			G_PayCnt = G_PayCnt + 1;
        //			functionReturnValue = "Y";
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			Error_Message:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			if (ErrNum == 1) {
        //				functionReturnValue = " 잠겨져 있습니다.";
        //			} else {
        //				functionReturnValue = Err().Description;
        //			}
        //			return functionReturnValue;
        //		}
        #endregion

        #region Get_ReAmt
        //		private double Get_ReAmt(ref string oJobDate, ref string oMstCode, ref string oSusil)
        //		{
        //			double functionReturnValue = 0;
        //			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			//반환컬럼,조건 컬럼,테이블,조건값,앤드절
        //			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			string sQry = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			sQry = "Exec PH_PY106  '" + Strings.Trim(oJobDate) + "', '" + Strings.Trim(oMstCode) + "', '" + oSusil + "'";
        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				functionReturnValue = 0;
        //			} else {
        //				functionReturnValue = Conversion.Val(oRecordSet.Fields.Item(0).Value);
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //		}
        #endregion

        #region Change_GOSIL
        //		private string Change_GOSIL(ref string xSuSil)
        //		{
        //			/// 공식안에 시스템제공코드있으면 해당 값으로 변경
        //			xSuSil = Strings.Replace(xSuSil, "X01", Convert.ToString(X01_Val));
        //			/// 기본일급
        //			xSuSil = Strings.Replace(xSuSil, "X02", Convert.ToString(X02_Val));
        //			/// 통상일급
        //			xSuSil = Strings.Replace(xSuSil, "X03", Convert.ToString(X03_Val));
        //			/// 기본시급
        //			xSuSil = Strings.Replace(xSuSil, "X04", Convert.ToString(X04_Val));
        //			/// 통상시급
        //			xSuSil = Strings.Replace(xSuSil, "X10", Convert.ToString(X10_Val));
        //			/// 월총일수
        //			xSuSil = Strings.Replace(xSuSil, "X11", Convert.ToString(X11_Val));
        //			/// 근속년수
        //			xSuSil = Strings.Replace(xSuSil, "X12", Convert.ToString(X12_Val));
        //			/// 근속월수
        //			xSuSil = Strings.Replace(xSuSil, "X13", Convert.ToString(X13_Val));
        //			/// 근속일수
        //			xSuSil = Strings.Replace(xSuSil, "X14", X14_Val);
        //			/// 당월입퇴사자유무
        //			xSuSil = Strings.Replace(xSuSil, "X16", Convert.ToString(X16_Val));
        //			/// 당월입퇴사근무기준일
        //			xSuSil = Strings.Replace(xSuSil, "X17", Convert.ToString(X17_Val));
        //			/// 수습율
        //			xSuSil = Strings.Replace(xSuSil, "X18", Convert.ToString(X18_Val));
        //			/// 수습일수
        //			xSuSil = Strings.Replace(xSuSil, "X19", Convert.ToString(X19_Val));
        //			/// 근속일수(급여종료일기준)
        //			xSuSil = Strings.Replace(xSuSil, "X20", Convert.ToString(X20_Val));
        //			/// 근속일수(상여종료일기준)
        //			xSuSil = Strings.Replace(xSuSil, "X21", "''" + oYM + "''");
        //			/// 귀속연월
        //			xSuSil = Strings.Replace(xSuSil, "X22", "''" + oJOBTYP + "''");
        //			/// 지급종류
        //			xSuSil = Strings.Replace(xSuSil, "X23", "''" + oJOBGBN + "''");
        //			/// 지급구분

        //			return xSuSil;
        //		}
        #endregion

        #region Raise_FormItemEvent
        //		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			string sQry = null;
        //			int i = 0;
        //			string JIGBIL = null;
        //			SAPbouiCOM.ComboBox oCombo = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;


        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					if (pval.BeforeAction == true) {

        //					} else if (pval.BeforeAction == false) {

        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					if (pval.BeforeAction == true & pval.ItemUID == "YM" & pval.CharPressed == 9 & pval.FormMode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					
        //					} else if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {

        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3

        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;

        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					if (pval.BeforeAction == false & pval.ItemChanged == true) {
        //						
        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					break;

        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
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

        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
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
        //					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //					//컬렉션에서 삭제및 모든 메모리 제거
        //					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //					if (pval.BeforeAction == true) {
        //					} else if (pval.BeforeAction == false) {

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

        //			} else if ((pval.BeforeAction == false)) {

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

        #endregion 백업소스코드_E
    }
}
