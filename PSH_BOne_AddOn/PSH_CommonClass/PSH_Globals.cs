using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    internal static class PSH_Globals
    {

        ////전역변수
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SBO_Application;

        //public static SAPbouiCOM.ProgressBar oProgBar;

        ////현재 폼의총갯수
        public static int FormCurrentCount;
        ////생성한 폼의총갯수
        public static int FormTotalCount;
        ////컬렉션 개체
        public static Collection ClassList;
        ////FormType 객체수
        public static int FormTypeListCount;
        ////FormType 객체
        public static Collection FormTypeList;

        //public static int SerialNo;

        //public static string oForm_ActiveItem;
        //public static short oForm_ActiveRow;

        ////Path/Srf/Rpt 패스
        ////XML메뉴경로
        public static string SP_XMLPath;
        ////PathINI경로
        public static string SP_Path;
        ////스크린폴더위치
        public static string Screen;
        ////레포트폴더위치
        public static string Report;

        ////ODBC
        //public static string SP_ODBC_YN;
        public static string SP_ODBC_IP; //서버 IP
        public static string SP_ODBC_Name;
        public static string SP_ODBC_DBName;
        public static string SP_ODBC_ID;
        public static string SP_ODBC_PW;

        ////Network Connection
        //public static string SP_NETWORK_YN;
        //public static string SP_NETWORK_DRIVE;
        //public static string SP_NETWORK_PATH;
        //public static string SP_NETWORK_ID;
        //public static string SP_NETWORK_PW;

        ////Cr부분
        //public static string ZG_CRWDSN;
        //public static ADODB.Connection g_ERPDMS;
        //public static ADODB.Recordset g_ADORS1;
        //public static ADODB.Recordset g_ADORS2;

        //public static CRAXDDRT.Application g_CApp;
        //public static CRAXDDRT.Report g_Report;
        //public static CRAXDDRT.FormulaFieldDefinition g_cFormula;
        //public static object g_GCrview;
        //public static CRAXDDRT.ParameterFieldDefinitions g_Params;
        //public static CRAXDDRT.ParameterFieldDefinition g_Param;

        //public static CRAXDDRT.Sections g_CrSections;
        //public static CRAXDDRT.Section g_CrSection;
        //public static CRAXDDRT.ReportObjects g_CrReportObjs;
        //public static CRAXDDRT.SubreportObject g_CrSubReportObj;
        //public static CRAXDDRT.Report g_CrSubReport;
        //public static CRAXDDRT.Database g_CrDB;

        //public static string[] gRpt_Formula;
        //public static string[] gRpt_Formula_Value;
        //public static string[] gRpt_Param;
        //public static string[] gRpt_Param_Value;
        //public static string[] gRpt_SRptSqry;
        //public static string[] gRpt_SRptName;
        //public static string[] gRpt_SFormula;
        //public static string[] gRpt_SFormula_Value;

        //public class ZPAY_g_EmpID
        //{
        //    public string EmpID; //사원순번
        //    public string MSTCOD;//사원번호
        //    public string MSTNAM; //사원성명
        //    public string TeamCode; //부서
        //    public string RspCode; //담당
        //    public string ClsCode; //반
        //    public string CLTCOD; //자사코드
        //    public string StartDate; //입사일자
        //    public string TermDate; //퇴사일자
        //    public string RETDAT; //퇴직정산일
        //    public string BALYMD; //최종발령일
        //    public string BALCOD; //최종부서
        //    public string JIGTYP; //직종
        //    public string Position; //직위
        //    public string JIGCOD; //직급
        //    public string HOBONG; //호봉
        //    public string PAYTYP; //급여형태
        //    public string PAYSEL; //급여지급일구분
        //    public short GONCNT; //공제인원
        //    public short DAGYSU; //다자녀추가공제
        //    public double STDAMT; //기본급
        //    public string GBHSEL; //고용보험여부
        //    public string PERNBR; //주민번호
        //    public string Sex; //성별
        //    public string GRPDAT; //그룹입사일
        //    public string ENDRET; //퇴직중간정산일
        //}


        ////사원조회 저장용 변수
        //public struct ZPAY_g_EmpID
        //{
        //    public string EmpID; //사원순번
        //    public string MSTCOD;//사원번호
        //    public string MSTNAM; //사원성명
        //    public string TeamCode; //부서
        //    public string RspCode; //담당
        //    public string ClsCode; //반
        //    public string CLTCOD; //자사코드
        //    public string StartDate; //입사일자
        //    public string TermDate; //퇴사일자
        //    public string RETDAT; //퇴직정산일
        //    public string BALYMD; //최종발령일
        //    public string BALCOD; //최종부서
        //    public string JIGTYP; //직종
        //    public string Position; //직위
        //    public string JIGCOD; //직급
        //    public string HOBONG; //호봉
        //    public string PAYTYP; //급여형태
        //    public string PAYSEL; //급여지급일구분
        //    public short GONCNT; //공제인원
        //    public short DAGYSU; //다자녀추가공제
        //    public double STDAMT; //기본급
        //    public string GBHSEL; //고용보험여부
        //    public string PERNBR; //주민번호
        //    public string Sex; //성별
        //    public string GRPDAT; //그룹입사일
        //    public string ENDRET; //퇴직중간정산일
        //}

        //1:근태, 2:급상여, 3:퇴직, 4:원천
        public static bool[] M_Used = new bool[5];
        //일근태사용
        //public static bool M_DayGNT;
        //년차사용
        //public static bool M_YunGNT;
        //개인별근태사용
        //public static bool M_PrsGNT;
        //정산기타소득사용
        //public static bool M_JsnGIT;
        //정산사업소득사용
        //public static bool M_JsnBUS;
        //정산이자소득사용
        //public static bool M_JsnEJA;
        //정산일용직사용
        //public static bool M_JsnILY;


        ////사용자구조체
        //public static string Value01;
        //public static string Value02;
        //public static string Value03;
        //public static string Value04;
        //public static string Value05;
        //public static string Value06;
        //public static string Value07;
        //public static string Value08;
        //public static string Value09;
        //public static string Value10;
        //public static string Value11;
        //public static string Value12;
        //public static string Value13;
        //public static string Value14;
        //public static string Value15;
        //public static string Value16;
        //public static string Value17;
        //public static string Value18;
        //public static string Value19;
        //public static string Value20;

        //public static int oTitleNameCount;

        //public static System.Windows.Forms.Form ZP_Form_Renamed;
        //public static System.Windows.Forms.Form frmRPT_View11;
        //public static System.Windows.Forms.Form frmRPT_View12;
        //public static System.Windows.Forms.Form frmRPT_View13;


        //근 속  년 수
        public static short ZPAY_GBL_GNSYER;
        //       월 수
        public static short ZPAY_GBL_GNSMON;
        //       일 수
        public static short ZPAY_GBL_GNSDAY;
        //근 무  년 수
        public static short ZPAY_GBL_GNMYER;
        //       월 수
        public static short ZPAY_GBL_GNMMON;
        //       일 수
        public static short ZPAY_GBL_GNMDAY;

        //정산년도
        //[VBFixedString(4)]
        //public static short ZPAY_GBL_JSNYER;

        /// <summary>
        /// SAP B1 이벤트 필터 자동 등록
        /// </summary>
        /// <param name="classType">Form Class</param>
        public static void ExecuteEventFilter(System.Type classType)
        {
            SAPbouiCOM.EventFilters eventFilters = new SAPbouiCOM.EventFilters();
            SAPbouiCOM.EventFilter eventFilter = null;

            eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK); //Main Menu 클릭 이벤트를 실행하기 위한 필수 이벤트(모든 클래스 필터 적용)

            try
            {
                System.Reflection.MethodInfo[] arrayMethodInfo = classType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly);

                for (int i = 0; i < arrayMethodInfo.Length; i++)
                {
                    System.Reflection.MethodInfo methodInfo = (System.Reflection.MethodInfo)arrayMethodInfo[i];

                    if (methodInfo.Name == "Raise_EVENT_ITEM_PRESSED")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_KEY_DOWN")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_GOT_FOCUS")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_LOST_FOCUS")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_COMBO_SELECT")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_CLICK")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_DOUBLE_CLICK")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_MATRIX_LINK_PRESSED")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_MATRIX_COLLAPSE_PRESSED")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_VALIDATE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_MATRIX_LOAD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_DATASOURCE_LOAD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_LOAD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_UNLOAD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_ACTIVATE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_DEACTIVATE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_CLOSE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_RESIZE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_KEY_DOWN")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_MENU_HILIGHT")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_PRINT")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_PRINT_DATA")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT_DATA);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_CHOOSE_FROM_LIST")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_RightClickEvent")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
                        eventFilter.AddEx(classType.Name);
                    }
                    //else if (methodInfo.Name == "Raise_EVENT_MENU_CLICK") //모든 클래스 필터 적용
                    //{
                    //    eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
                    //    eventFilter.AddEx(classType.Name);
                    //}
                    else if (methodInfo.Name == "Raise_EVENT_FORM_DATA_ADD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_DATA_UPDATE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_DATA_DELETE")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE);
                        eventFilter.AddEx(classType.Name);
                    }
                    else if (methodInfo.Name == "Raise_EVENT_FORM_DATA_LOAD")
                    {
                        eventFilter = eventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
                        eventFilter.AddEx(classType.Name);
                    }
                }

                SBO_Application.SetFilter(eventFilters);
            }
            catch(System.Exception ex)
            {
                SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(eventFilter);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(eventFilters);
            }
        }
    }

    public class ZPAY_g_EmpID
    {
        public string EmpID; //사원순번
        public string MSTCOD;//사원번호
        public string MSTNAM; //사원성명
        public string TeamCode; //부서
        public string RspCode; //담당
        public string ClsCode; //반
        public string CLTCOD; //자사코드
        public string StartDate; //입사일자
        public string TermDate; //퇴사일자
        public string RETDAT; //퇴직정산일
        //public string BALYMD; //최종발령일
        //public string BALCOD; //최종부서
        public string JIGTYP; //직종
        public string Position; //직위
        public string JIGCOD; //직급
        public string HOBONG; //호봉
        public string PAYTYP; //급여형태
        public string PAYSEL; //급여지급일구분
        public short GONCNT; //공제인원
        public short DAGYSU; //다자녀추가공제
        public double STDAMT; //기본급
        public string GBHSEL; //고용보험여부
        public string PERNBR; //주민번호
        public string Sex; //성별
        public string GRPDAT; //그룹입사일
        public string ENDRET; //퇴직중간정산일
    }
}
