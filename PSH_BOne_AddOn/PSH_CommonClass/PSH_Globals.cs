using System;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    internal static class PSH_Globals
    {
        public static SAPbobsCOM.Company oCompany; //전역변수
        public static SAPbouiCOM.Application SBO_Application;
        public static int FormCurrentCount; //현재 폼의총갯수
        public static int FormTotalCount; //생성한 폼의총갯수
        public static Collection ClassList; //컬렉션 개체
        public static int FormTypeListCount; //FormType 객체수
        public static Collection FormTypeList; //FormType 객체

        public static string SP_XMLPath; //XML메뉴경로
        public static string SP_Path; //PathINI경로
        public static string Screen; //스크린폴더위치
        public static string Report; //레포트폴더위치

        //ODBC
        public static string SP_ODBC_IP; //서버 IP
        public static string SP_ODBC_Name;
        public static string SP_ODBC_DBName;
        public static string SP_ODBC_ID;
        public static string SP_ODBC_PW;
        
        public static bool[] M_Used = new bool[5]; //1:근태, 2:급상여, 3:퇴직, 4:원천
        public static short ZPAY_GBL_GNSYER; //근속년수
        public static short ZPAY_GBL_GNSMON; //월수
        public static short ZPAY_GBL_GNSDAY; //일수
        public static short ZPAY_GBL_GNMYER; //근무년수
        public static short ZPAY_GBL_GNMMON; //월수
        public static short ZPAY_GBL_GNMDAY; //일 수
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
