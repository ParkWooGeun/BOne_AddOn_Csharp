//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.Globalization;
//using System.IO;
//using System.Linq;
//using System.Reflection;
//using System.Runtime.CompilerServices;
//using System.Security;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.VisualBasic;

//[System.Runtime.InteropServices.ProgId("ZPY521_NET.ZPY521")]
//public class ZPY521
//{
//    // //  SAP MANAGE UI API 2004 SDK Sample
//    // //****************************************************************************
//    // //  File           : ZPY521.cls
//    // //  Module         : 인사관리>정산관리
//    // //  Desc           : 근로소득전산매체수록
//    // //  FormType       : 2000060521
//    // //  Create Date    : 2006.01.24
//    // //  Modified Date  :
//    // //  Creator        : Ham Mi Kyoung
//    // //  Modifier       :
//    // //  Copyright  (c) Morning Data
//    // //****************************************************************************

//    public string oFormUniqueID;
//    public SAPbouiCOM.Form oForm;
//    private SAPbobsCOM.Recordset sRecordset;
//    private SAPbouiCOM.Matrix oMat1;
//    private string Last_Item; // 클래스에서 선택한 마지막 아이템 Uid값

//    private string oJsnYear;
//    private string JSNGBN;
//    private string STRMON;
//    private string ENDMON;
//    private string MSTBRK;
//    private string DPTSTR;
//    private string DPTEND;
//    private string MSTCOD;
//    private string oFilePath;

//    private VB6.FixedLengthString FILNAM = new VB6.FixedLengthString(30); // 파  일  명
//    private int MaxRow;
//    private short BUSCNT; // / B레코드일련번호
//    private short BUSTOT; // / B레코드총갯수

//    private short NEWCNT;
//    private short OLDCNT;
//    private string C_MSTCOD;
//    private string C_CLTCOD;
//    private string E_BUYCNT;
//    private string C_BUYCNT;
//    private string CLTCOD;

//    private struct Arecord
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PRTDAT; // 제출  일자
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RPTGBN; // 제  출  자 (1;세무대리인, 2;법인, 3;개인)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] TAXAGE; // 세무대리인
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] HOMTID; // 홈텍스ID
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] PGMCOD; // 세무프로그램코드
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(40)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//        public char[] SANGHO; // 상      호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] DAMDPT; // 담당부서명
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] DAMNAM; // 담당 성명
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(15)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 15)]
//        public char[] DAMTEL; // 담당전화번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] BUSCNT; // B Record수(신고의무자수)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] HANCOD; // 한글코드종류
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1082)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1082)]
//        public char[] FILLER; // 공      란
//    }
//    private Arecord Arec;

//    private struct Brecord
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] BUSCNT; // 일련  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(40)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//        public char[] SANGHO; // 상      호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] COMPRT; // 대  표  자
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] PERNBR; // 주민  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(7)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//        public char[] NEWCNT; // C Record수
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(7)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//        public char[] OLDCNT; // D Record수
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(14)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
//        public char[] INCOME; // 소득금액총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] GULGAB; // 결정소득세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] GULCOM; // 법인세(공란)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] GULJUM; // 주민세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] GULNON; // 농특세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] GULTOT; // 총  액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RNGCOD; // 제출대상기간
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1061)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1061)]
//        public char[] FILLER; // 공      란
//    }
//    private Brecord Brec;

//    // / 근로 주(현) 근무처 레코드 /
//    private struct Crecord
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] SQNNBR; // 일련  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] JONCNT; // 종전근무처수
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] DWEGBN; // 거주자구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] RGNCOD; // 거주지국코드
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RGNTAX; // 외국인단일세율적용
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] MSTNAM; // 성      명
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] INTGBN; // 내외국인구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] PERNBR; // 주민  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] GUKCOD; // 국적코드                                  (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] HUSMAN; // 세대주여부                                (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] JSNGBN; // 연말정산구분                              (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNB1; // 주(현) 사업자번호                         (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(40)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//        public char[] SANGHO; // 주(현) 근무처명                           (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] STRINT; // 귀속년도시작=>근무기간시작연월일
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] ENDINT; // 귀속년도종료=>근무기간종료연월일
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] STRGAM; // 감면기간시작
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] ENDGAM; // 감면기간종료101
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] PAYAMT; // 급여  총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] BNSAMT; // 상여  총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] INJBNS; // 인정  상여
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] JUSBNS; // 주식매수선택권행사이익                    (2007년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] URIBNS; // 우리사주조합인출금                        (2009년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] FILD01; // 공란(2009년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] TOTAMT; // 급상여총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGG01; // 비 과 세(G01:학자금)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH01; // 비 과 세(H01:무보수위원수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH05; // 비 과 세(H05:경호,승선수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH06; // 비 과 세(H06:유아,초중등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH07; // 비 과 세(H07:고등교육법)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH08; // 비 과 세(H08:특별법)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH09; // 비 과 세(H09:연구기관등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH10; // 비 과 세(H10:기업연구소)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH11; // 비 과 세(H11:취재수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH12; // 비 과 세(H12:벽지수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH13; // 비 과 세(H13:재해관련급여)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGI01; // 비 과 세(I01:외국정부등근로자)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGK01; // 비 과 세(K01:외국주둔군인등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM01; // 비 과 세(M01:국외근로100만원)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM02; // 비 과 세(M02:국외근로150만원)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM03; // 비 과 세(M03:국외근로)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGO01; // 비 과 세(O01:야간근로수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGQ01; // 비 과 세(Q01:출산보육수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGR10; // 비 과 세(R10:근로장학금)                  (2011년 신설)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGS01; // 비 과 세(S01:주식매수선택권)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGT01; // 비 과 세(T01:외국인기술자)
//                              // BIGX01 As String * 10    '비 과 세(X01:외국인근로자)                (2010년 폐지)
//                              // BIGY01 As String * 10    '비 과 세(Y01:우리사주조합배정)            (2010년 폐지)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY02; // 비 과 세(Y02:우리사주조합인출금50%)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY03; // 비 과 세(Y03:우리사주조합인출금75%)
//                              // BIGY20 As String * 10    '비 과 세(Y20:주택자금보조금)              (2010년 폐지)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY21; // 비 과 세(Y21:장기미취업자 중소기업 취업)  (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGZ01; // 비 과 세(Z01:해저광물자원개발)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY22; // 비 과 세(그밖의비과세)=명확화:(Y22:지정비과세)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGTOT; // 비과세 합계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGTO1; // 감면소득 계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] INCOME; // 근로소득수입금액90
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILGNL; // 근로소득공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] GNLOSD; // 근로소득금액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJBAS; // 본인  공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJBWO; // 배우자공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] BUYNSU; // 부양가족인원
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJBYN; // 부양가족공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] GYNGLO; // 경로우대인원
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJGYN; // 경로우대공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] JANGAE; // 장애자인원
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJJAE; // 장애자공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJBNY; // 부녀자공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] BUYN06; // 자녀양육공제인원
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJSON; // 자녀양육공제금액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] CHLSAN; // 출산입양자공제인원
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJCHL; // 출산입양자공제금액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] FILD02; // 공란
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DAGYSU; // 다자녀추가공제인원                        (2007년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] INJADD; // 소수추가공제->2007년부터 다자녀추가공제로변경
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] KUKGON; // 국민연금보험료공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] YUNGON; // 기타연금보험료공제(공무원연금)            (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] YUNGO1; // 기타연금보험료공제(군인연금)              (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] YUNGO2; // 기타연금보험료공제(사립학교교직원연금)    (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] YUNGO3; // 기타연금보험료공제(별정우체국연금)        (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITRE2; // 퇴직연금소득공제(과학기술인공제)          (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITRET; // 퇴직연금소득공제(근로자퇴직급여보장법)    (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILMBH; // 보험료_건강보험료                         (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILGBH; // 보험료_고용보험료                         (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILBHM; // 보험료_보장성보험                         (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILJHM; // 보험료_장애인전용                         (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILMED; // 의료비공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILSCH; // 교육비공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILHUS; // 주택자금_주택임대차차입금 원리금상환공제금액-대출기관(구:주택자금공제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILHU2; // 주택자금_주택임대차차입금 원리금상환공제금액-거주자(구:주택자금공제) (2011년)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILWOL; // 주택자금_월세액                           (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILJHE; // 장기주택저당차입금이자상환공제금액-15년미만(2011년 년수별세분화됨)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILJH2; // 장기주택저당차입금이자상환공제금액-29년미만(2011년)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILJH3; // 장기주택저당차입금이자상환공제금액-30년이상(2011년)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILGBU; // 기부금공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] PILFLD; // 공란
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] PILTOT; // 계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] PILGON; // 표준공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] CHAGAM; // 차감소득금액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] GITGYN; // 개인연금저축소득공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] GITYUN; // 연금소득공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITSGI; // 소기업공제부금소득공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITHUS; // 주택마련저축소득공제_청약저축             (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITHU1; // 주택마련저축소득공제_주택청약종합저축     (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITHU2; // 주택마련저축소득공제_장기주택마련저축     (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITHU3; // 주택마련저축소득공제_근로자주택마련저축   (2010년 항목분리)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITINV; // 투자조합공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] GITCAD; // 신용카드공제
//                              // GITUSG As String * 1     '우리사주조합소득공제(음수1 양수0)         (2010년 삭제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITUSJ; // 우리사주조합출연금
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITJFD; // 장기주식형저축소득공제
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITGYU; // 고용유지중소기업근로자소득공제(2009년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] FILD03; // 공란(2009년추가)
//                              // GITTOG As String * 1     '기타소득공제계(기호 음수1, 양수0)         (2010년 삭제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GITTOT; // 기타소득공제계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] TAXSTD; // 종합소득과세표준
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] SANTAX; // 산출  세액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GAMSOD; // 소득세  법
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GAMJOS; // 조  감  법
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GAMJYK; // 조 세 조약(2011년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GAMFLD; // 공      란
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GAMTOT; // 세액감면계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] TAXGNL; // 근로  소득
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NABSEE; // 납세  조합
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] TAXBRO; // 주택  차입
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] TAXGBU; // 기부정치자금
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] TAXFRG; // 외국  납부
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] FILD04; // 공란
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] TAXTOT; // 세액공제계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GULGAB; // 결정소득세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GULJUM; // 주민세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] GULNON; // 특별세
//                              // GULTOT As String * 10    '    세액계(2011년삭제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANGAB; // 현기납부소득세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANJUM; // 주민세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANNON; // 특별세
//                              // NANTOT As String * 10    '        세액계(2011년 삭제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] CHAGAG; // 현)차감소득세(기호 음수1, 양수0)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] CHAGAB; // 현)    소득세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] CHAJUG; // 현)    주민세(기호 음수1, 양수0)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] CHAJUM; // 주민세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] CHANOG; // 현)    특별세(기호 음수1, 양수0)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] CHANON; // 특별세
//                              // CHATOG As String * 1     '현)    세액계(기호 음수1, 양수0)(2011년삭제)
//                              // CHATOT As String * 10    '       세액계(2011년삭제)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] FILLER; // 공      란
//    }
//    private Crecord Crec;

//    private struct D_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] SQNNBR; // 일련  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(50)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 50)]
//        public char[] FILD01; // 공      란
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] PERNBR; // 주민  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] TAXJOH; // 납세조합구분                               (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(40)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//        public char[] JONNAM; // 근무처  명
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] JONNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] STRINT; // 근무기간시작연월일                         (2009년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] ENDINT; // 근무기간종료연월일                         (2009년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] STRGAM; // 감면기간시작                               (2009년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] ENDGAM; // 감면기간종료                               (2009년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] PAYAMT; // 급여  총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] BNSAMT; // 상여  총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] INJBNS; // 인정  상여
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] JUSBNS; // 주식매수선택권행사이익                     (2007년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] URIBNS; // 우리사주조합인출금                         (2009년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] FILD02; // 공란                                       (2009년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] TOTAMT; // 급상여총액
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGG01; // 비 과 세(G01:학자금)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH01; // 비 과 세(H01:무보수위원수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH05; // 비 과 세(H05:경호,승선수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH06; // 비 과 세(H06:유아,초중등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH07; // 비 과 세(H07:고등교육법)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH08; // 비 과 세(H08:특별법)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH09; // 비 과 세(H09:연구기관등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH10; // 비 과 세(H10:기업연구소)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH11; // 비 과 세(H11:취재수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH12; // 비 과 세(H12:벽지수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGH13; // 비 과 세(H13:재해관련급여)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGI01; // 비 과 세(I01:외국정부등근로자)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGK01; // 비 과 세(K01:외국주둔군인등)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM01; // 비 과 세(M01:국외근로100만원)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM02; // 비 과 세(M02:국외근로150만원)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGM03; // 비 과 세(M02:국외근로)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGO01; // 비 과 세(O01:야간근로수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGQ01; // 비 과 세(Q01:출산보육수당)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGR10; // 비 과 세(R10:근로장학금)                   (2011년추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGS01; // 비 과 세(S01:주식매수선택권)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGT01; // 비 과 세(T01:외국인기술자)
//                              // BIGX01 As String * 10    '비 과 세(X01:외국인근로자)                (2010년 폐지)
//                              // BIGY01 As String * 10    '비 과 세(Y01:우리사주조합배정)            (2010년 폐지)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY02; // 비 과 세(Y02:우리사주조합인출금50%)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY03; // 비 과 세(Y03:우리사주조합인출금75%)
//                              // BIGY20 As String * 10    '비 과 세(Y20:주택자금보조금)              (2010년 폐지)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY21; // 비 과 세(Y21:장기미취업자 중소기업취업)   (2010년 추가)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGZ01; // 비 과 세(Z01:해저광물자원개발)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGY22; // 비 과 세(그밖의비과세->지정비과세)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGTOT; // 비과세합계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BIGTO1; // 감면소득 계
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANGAB; // 현기납부소득세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANJUM; // 주민세
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] NANNON; // 특별세
//                              // NANTOT As String * 10    '       세액계 (2011년 폐지)
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] JONCNT; // 종전근무처일련번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(692)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 692)]
//        public char[] FILLER; // 공      란
//    }
//    private D_Record Drec;

//    // / 부양가족공제자명세 레코드 /
//    private struct E_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] SQNNBR; // 일련  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] PERNBR; // 주민  번호
//                              // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKCOD = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 관계코드
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKINT = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 내외국인구분
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKNAM = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 성명
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKPER = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 주민등록번호
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKBAS = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기본공제
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKJAN = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 장애자공제
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKBY6 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 자녀양육
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKBUY = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 부녀자             (2007년추가)
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKJEL = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 경로우대           (2007년추가)
//                                                                                                                                          // CHKDAG(1 To 5) As String * 1    '다자녀(2007년추가) ->2009년제외
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CHKCHS = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 출산입양           (2008년추가)
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] BOHAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 보험료
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] MEDAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 의료비
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] EDCAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 교육비
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CADAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 신용카드
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CSHCA1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 직불카드           (2010년 추가)
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CSHAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 현금영수증
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] GBUAM1 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기부금(2009년추가)
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] BOHAM2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 보험료
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] MEDAM2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 의료비
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] EDCAM2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 교육비
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CADAM2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 신용카드
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] CSHCA2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 직불카드           (2010년 추가)
//                                                                                                                                          // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] GBUAM2 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기부금
//                                                                                                                                          // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] FAMNBR; // 부양가족일련번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(368)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 368)]
//        public char[] FILLER; // 공      란
//    }
//    // UPGRADE_WARNING: Erec 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//    private E_Record Erec;

//    private struct F_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] RECGBN; // 레코드구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] DTAGBN; // 자료  구분
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] TAXCOD; // 세  무  서
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] SQNNBR; // 일련  번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] BUSNBR; // 사업자번호
//                              // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] PERNBR; // 주민  번호
//                              // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SAVGBN = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 소득공제구분
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SAVCOD = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 금융기관코드
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SAVNAM = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 금융기관상호
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SAVNUM = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 계좌번호
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] STYEAR = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 납입연차
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SAVAMT = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 불입금액
//                                                                                                                                           // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] SARAMT = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 공제금액
//                                                                                                                                           // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(70)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 70)]
//        public char[] FILLER;
//    }
//    // UPGRADE_WARNING: Frec 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//    private F_Record Frec;

//    // *******************************************************************
//    // .srf 파일로부터 폼을 로드한다.
//    // *******************************************************************
//    public void LoadForm()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo LoadForm_Error' at character 95139
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo LoadForm_Error

// */		int i;
//        MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//        oXmlDoc.Load(MDC_Globals.SP_Path + @"\" + SP_Screen + @"\ZPY521.srf");
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount);

//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        // //여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10);
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10);

//        Sbo_Application.LoadBatchActions((oXmlDoc.xml));

//        oFormUniqueID = "ZPY521_" + GetTotalFormsCount;

//        // 폼 할당
//        oForm = Sbo_Application.Forms.Item(oFormUniqueID);

//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        // 컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        AddForms(this, oFormUniqueID, "ZPY521");
//        oForm.SupportedModes = -1;
//        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//        oForm.Freeze(true);
//        CreateItems();
//        oForm.Freeze(false);

//        oForm.EnableMenu(("1281"), false); // / 찾기
//        oForm.EnableMenu(("1282"), true); // / 추가
//        oForm.EnableMenu(("1284"), false); // / 취소
//        oForm.EnableMenu(("1293"), false); // / 행삭제
//        oForm.Update();
//        oForm.Visible = true;

//        // UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oXmlDoc = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        LoadForm_Error:
//        ;

//        // UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oXmlDoc = null/* TODO Change to default(_) if this is not a reference type */;
//        Sbo_Application.StatusBar.SetText("Form_Load Error:" + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        if ((oForm == null) == false)
//        {
//            oForm.Freeze(false);
//            // UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oForm = null;
//        }
//    }
//    // *******************************************************************
//    // // ItemEventHander
//    // *******************************************************************
//    public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//    {
//        string sQry;
//        int i;
//        SAPbouiCOM.ComboBox oCombo;
//        SAPbouiCOM.Column oColumn;
//        SAPbouiCOM.Columns oColumns;
//        SAPbobsCOM.Recordset oRecordSet;
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Raise_FormIte...' at character 98143
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
		
//		On Error GoTo Raise_FormItemEvent_Error

// */
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        switch (pval.EventType)
//        {
//            case object _ when SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                {
//                    if (pval.BeforeAction)
//                    {
//                        if (pval.ItemUid == "CBtn1")
//                        {
//                            oForm.Items.Item("MSTCOD").CLICK(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            Sbo_Application.ActivateMenuItem(("7425"));
//                            BubbleEvent = false;
//                        }
//                        else if (pval.ItemUid == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                        {
//                            if (HeaderSpaceLineDel == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            if (File_Create == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            else
//                            {
//                                BubbleEvent = false;
//                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//                            }
//                        }
//                        else if (pval.ItemUid == "Btn1")
//                        {
//                            oFilePath = ZP_Form.vbGetBrowseDirectory(ZP_Form);
//                            oForm.DataSources.UserDataSources.Item("Path").ValueEx = oFilePath;
//                            BubbleEvent = false;
//                            return;
//                        }
//                    }
//                    else
//                    {
//                    }

//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                {
//                    if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUid == "JsnYear" | pval.ItemUid == "MSTCOD"))
//                        FlushToItemValue(pval.ItemUid);
//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_CLICK:
//                {
//                    if (pval.BeforeAction == true & pval.ItemUid != "1000001" & pval.ItemUid != "2")
//                    {
//                        if (Last_Item == "JsnYear")
//                        {
//                            // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            if (Trim(oForm.Items.Item(Last_Item).Specific.Value) != "")
//                            {
//                                // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                if (MDC_SetMod.ChkYearMonth(Trim(System.Convert.ToString(oForm.Items.Item(Last_Item).Specific.Value)) + "01") == false)
//                                {
//                                    oForm.Items.Item(Last_Item).Update();
//                                    Sbo_Application.StatusBar.SetText("정산 연도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                                    BubbleEvent = false;
//                                }
//                            }
//                        }
//                        else if (Last_Item == "MSTCOD")
//                        {
//                            // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            if (Trim(oForm.Items.Item(Last_Item).Specific.String) != "" & MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Trim(oForm.Items.Item(Last_Item).Specific.String) + "'", "") == true)
//                            {
//                                oForm.Items.Item(Last_Item).Update();
//                                Sbo_Application.StatusBar.SetText("사원 번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                                BubbleEvent = false;
//                            }
//                        }
//                    }

//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                {
//                    if (pval.BeforeAction == true & pval.ItemUid == "JsnYear" & pval.CharPressed == 9)
//                    {
//                        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (Len(Trim(oForm.Items.Item(pval.ItemUid).Specific.String)) < 4)
//                            // UPGRADE_WARNING: oForm.Items(pval.ItemUid).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item(pval.ItemUid).Specific.Value = VB6.Format(oForm.Items.Item(pval.ItemUid).Specific.Value, "2000");
//                        // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (MDC_SetMod.ChkYearMonth(Trim(System.Convert.ToString(oForm.Items.Item(pval.ItemUid).Specific.Value)) + "01") == false)
//                        {
//                            Sbo_Application.StatusBar.SetText("정산 연도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                            BubbleEvent = false;
//                        }
//                    }
//                    else if (pval.BeforeAction == true & pval.ItemUid == "MSTCOD" & pval.CharPressed == 9)
//                    {
//                        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (Trim(oForm.Items.Item("MSTCOD").Specific.String) != "" & MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", "") == true)
//                        {
//                            Sbo_Application.StatusBar.SetText("사원 번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                            BubbleEvent = false;
//                        }
//                    }

//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                {
//                    if (Last_Item == "Mat1")
//                    {
//                        if (pval.Row > 0)
//                            Last_Item = pval.ItemUid;
//                    }
//                    else
//                        Last_Item = pval.ItemUid;
//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                {
//                    // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    // 컬렉션에서 삭제및 모든 메모리 제거
//                    // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    if (pval.BeforeAction == false)
//                    {
//                        RemoveForms(oFormUniqueID);
//                        // UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oForm = null;
//                        // UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oMat1 = null;
//                    }

//                    break;
//                }
//        }

//        return;
//        // /////////////////////////////////////////////////////////////////////////////////////////////////////////////
//        Raise_FormItemEvent_Error:
//        ;
//        Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:" + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//    }

//    // *******************************************************************
//    // // MenuEventHander
//    // *******************************************************************
//    public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//    {
//        if (pval.BeforeAction == true)
//            return;

//        switch (pval.MenuUID)
//        {
//            case "1287" // / 복제
//           :
//                {
//                    break;
//                }

//            case "1281":
//            case "1282":
//                {
//                    oForm.Items.Item("JsnYear").CLICK(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    break;
//                }

//            case object _ when "1288" <= pval.MenuUID && pval.MenuUID <= "1291":
//                {
//                    break;
//                }

//            case "1293":
//                {
//                    break;
//                }
//        }
//        return;
//    }

//    public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//    {
//        int i;
//        string sQry;
//        SAPbouiCOM.ComboBox oCombo;

//        SAPbobsCOM.Recordset oRecordSet;
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Raise_FormDat...' at character 105852
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
		
		
//		On Error GoTo Raise_FormDataEvent_Error

// */
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        if ((BusinessObjectInfo.BeforeAction == false))
//        {
//            switch (BusinessObjectInfo.EventType)
//            {
//                case object _ when SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD // //33
//               :
//                    {
//                        break;
//                    }

//                case object _ when SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD // //34
//       :
//                    {
//                        break;
//                    }

//                case object _ when SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE // //35
//       :
//                    {
//                        break;
//                    }

//                case object _ when SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE // //36
//       :
//                    {
//                        break;
//                    }
//            }
//        }
//        // UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;

//        Raise_FormDataEvent_Error:
//        ;

//        // UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Information.Err.Number + " - " + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//    }

//    private void CreateItems()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 107359
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		SAPbouiCOM.ComboBox oCombo1;
//        SAPbouiCOM.ComboBox oCombo2;
//        SAPbobsCOM.Recordset oRecordSet;
//        SAPbouiCOM.EditText oEdit;
//        SAPbouiCOM.Column oColumn;
//        string sQry;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        oForm.DataSources.UserDataSources.Add("JsnYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4); // / 생성년도
//        oForm.DataSources.UserDataSources.Add("JSNGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10); // / 생성구분
//        oForm.DataSources.UserDataSources.Add("JSNTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10); // / 생성구분
//        oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10); // / 지점
//        oForm.DataSources.UserDataSources.Add("DptStr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10); // / 부서코드
//        oForm.DataSources.UserDataSources.Add("DptEnd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//        oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//        oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//        oForm.DataSources.UserDataSources.Add("EmpID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//        oForm.DataSources.UserDataSources.Add("PRTDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//        oForm.DataSources.UserDataSources.Add("SMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//        oForm.DataSources.UserDataSources.Add("EMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
//        oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//        oEdit = oForm.Items.Item("JsnYear").Specific;
//        oEdit.DataBind.SetBound(true, "", "JsnYear");
//        oEdit = oForm.Items.Item("MSTCOD").Specific;
//        oEdit.DataBind.SetBound(true, "", "MSTCOD");
//        oEdit = oForm.Items.Item("MSTNAM").Specific;
//        oEdit.DataBind.SetBound(true, "", "MSTNAM");
//        oEdit = oForm.Items.Item("EmpID").Specific;
//        oEdit.DataBind.SetBound(true, "", "EmpID");
//        oEdit = oForm.Items.Item("Path").Specific;
//        oEdit.DataBind.SetBound(true, "", "Path");
//        oEdit = oForm.Items.Item("PRTDAT").Specific;
//        oEdit.DataBind.SetBound(true, "", "PRTDAT");
//        oEdit = oForm.Items.Item("SMonth").Specific;
//        oEdit.DataBind.SetBound(true, "", "SMonth");
//        oEdit = oForm.Items.Item("EMonth").Specific;
//        oEdit.DataBind.SetBound(true, "", "EMonth");

//        // // 생성구분
//        oCombo1 = oForm.Items.Item("JSNGBN").Specific;

//        oCombo1.ValidValues.Add("1", "연말정산(재직자)");
//        oCombo1.ValidValues.Add("2", "중도정산(퇴사자)");
//        oCombo1.ValidValues.Add("%", "전체");

//        oCombo1.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue); // / 전체

//        // // 생성구분
//        oCombo1 = oForm.Items.Item("JSNTYP").Specific;
//        oCombo1.ValidValues.Add("1", "연간(01.01~12.31)지급분");
//        oCombo1.ValidValues.Add("2", "폐업에 의한 수시 제출분");
//        oCombo1.ValidValues.Add("3", "수시 분할제출분");
//        oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); // / 전체

//        // // 신고의무자
//        oCombo1 = oForm.Items.Item("CLTCOD").Specific;
//        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//        // sQry = " SELECT T0.U_WCHCLT, MAX(T1.NAME)  FROM [@PH_PY005A] T0 INNER JOIN [@PH_PY005A] T1 ON T0.U_WCHCLT = T1.CODE GROUP BY T0.U_WCHCLT  ORDER BY T0.U_WCHCLT"
//        oRecordSet.DoQuery(sQry);
//        while (!oRecordSet.EOF)
//        {
//            oCombo1.ValidValues.Add(Trim(oRecordSet.Fields.Item(0).Value), Trim(oRecordSet.Fields.Item(1).Value));
//            oRecordSet.MoveNext();
//        }
//        if (oCombo1.ValidValues.Count > 0)
//            oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);// / 전체

//        // // 사업장
//        oCombo1 = oForm.Items.Item("BPLId").Specific;
//        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//        oRecordSet.DoQuery(sQry);
//        oCombo1.ValidValues.Add("%", "모두");
//        while (!oRecordSet.EOF)
//        {
//            oCombo1.ValidValues.Add(Trim(oRecordSet.Fields.Item(0).Value), Trim(oRecordSet.Fields.Item(1).Value));
//            oRecordSet.MoveNext();
//        }
//        if (oCombo1.ValidValues.Count > 0)
//            oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);// / 전체
//                                                                // // 부서
//        oCombo1 = oForm.Items.Item("DptStr").Specific;
//        oCombo2 = oForm.Items.Item("DptEnd").Specific;
//        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'";
//        oRecordSet.DoQuery(sQry);
//        oCombo1.ValidValues.Add("-1", "모두");
//        oCombo2.ValidValues.Add("-1", "모두");
//        while (!oRecordSet.EOF)
//        {
//            oCombo1.ValidValues.Add(Trim(oRecordSet.Fields.Item(0).Value), Trim(oRecordSet.Fields.Item(1).Value));
//            oCombo2.ValidValues.Add(Trim(oRecordSet.Fields.Item(0).Value), Trim(oRecordSet.Fields.Item(1).Value));
//            oRecordSet.MoveNext();
//        }
//        if (oCombo1.ValidValues.Count > 0)
//        {
//            oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//            oCombo2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//        }
//        oMat1 = oForm.Items.Item("Mat1").Specific;

//        oForm.DataSources.UserDataSources.Add("Col0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//        oForm.DataSources.UserDataSources.Add("Col1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//        oColumn = oMat1.Columns.Item("Col0");
//        oColumn.DataBind.SetBound(true, "", "Col0");

//        oColumn = oMat1.Columns.Item("Col1");
//        oColumn.DataBind.SetBound(true, "", "Col1");

//        if (Trim(ZPAY_GBL_JSNYER.Value) == "" | Trim(ZPAY_GBL_JSNYER.Value) == "*")
//            oForm.DataSources.UserDataSources.Item("JsnYear").ValueEx = VB6.Format(DateTime.Now, "YYYY");
//        else
//            oForm.DataSources.UserDataSources.Item("JsnYear").ValueEx = ZPAY_GBL_JSNYER.Value;
//        if (Trim(ZPAY_GBL_JSNYER.Value) != "")
//        {
//            oForm.DataSources.UserDataSources.Item("SMonth").ValueEx = "01";
//            oForm.DataSources.UserDataSources.Item("EMonth").ValueEx = "12";
//        }


//        // UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oEdit = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo1 = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo2 = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oEdit = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo1 = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oCombo2 = null/* TODO Change to default(_) if this is not a reference type */;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//    }

//    private bool File_Create()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 114869
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        string oStr;
//        string sQry;

//        sRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        ErrNum = 0;
//        // / Question
//        if (Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }

//        oMat1.Clear();
//        MaxRow = 0;

//        BUSCNT = 0; // / B레크드 일련번호
//        BUSTOT = 0;
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        JSNGBN = oForm.Items.Item("JSNGBN").Specific.Selected.Value;
//        STRMON = VB6.Format(oForm.DataSources.UserDataSources.Item("SMonth").ValueEx, "00");
//        ENDMON = VB6.Format(oForm.DataSources.UserDataSources.Item("EMonth").ValueEx, "00");
//        // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oJsnYear = oForm.Items.Item("JsnYear").Specific.Value;
//        // / 파일경로설정
//        if (oFilePath == "")
//            oFilePath = @"C:\EOSDATA";
//        oFilePath = Interaction.IIf(Strings.Right(oFilePath, 1) == @"\", oFilePath, oFilePath + @"\");
//        oStr = CreateFolder(Strings.Trim(oFilePath));
//        if (Strings.Trim(oStr) != "")
//        {
//            ErrNum = 5;
//            goto Error_Message;
//        }

//        // / 근로 제출자(대리인) 레코드
//        if (File_Create_ARecord == false)
//        {
//            ErrNum = 2;
//            goto Error_Message;
//        }

//        FileSystem.FileClose(1);
//        FileOpen(1, FILNAM.Value, OpenMode.Output);
//        // / A레코드: 근로 원천징수의무자별 집계 레코드
//        PrintLine(1, MDC_SetMod.sStr(Arec.RECGBN) + MDC_SetMod.sStr(Arec.DTAGBN) + MDC_SetMod.sStr(Arec.TAXCOD) + MDC_SetMod.sStr(Arec.PRTDAT) + MDC_SetMod.sStr(Arec.RPTGBN) + MDC_SetMod.sStr(Arec.TAXAGE) + MDC_SetMod.sStr(Arec.HOMTID) + MDC_SetMod.sStr(Arec.PGMCOD) + MDC_SetMod.sStr(Arec.BUSNBR) + MDC_SetMod.sStr(Arec.SANGHO) + MDC_SetMod.sStr(Arec.DAMDPT) + MDC_SetMod.sStr(Arec.DAMNAM) + MDC_SetMod.sStr(Arec.DAMTEL) + MDC_SetMod.sStr(Arec.BUSCNT) + MDC_SetMod.sStr(Arec.HANCOD) + MDC_SetMod.sStr(Arec.FILLER));


//        Matrix_AddRow("제출자 레코드 생성 완료!", ref true);

//        // / B레코드: 근로 집계 레코드 /***********************************************/
//        sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] ";
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        sQry = sQry + " WHERE  U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "' ORDER BY Code";
//        sRecordset.DoQuery(sQry);
//        while (!sRecordset.EOF)
//        {
//            // / B레코드: 근로 집계 레코드 /***********************************************/
//            // UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Brec.TAXCOD = sRecordset.Fields.Item("U_TAXCODE").Value;
//            Brec.BUSNBR = Replace(sRecordset.Fields.Item("U_BUSNUM").Value, "-", "");
//            Brec.SANGHO = Trim(sRecordset.Fields.Item("U_CLTNAME").Value);
//            Brec.COMPRT = Trim(sRecordset.Fields.Item("U_COMPRT").Value);
//            Brec.PERNBR = Replace(Trim(sRecordset.Fields.Item("U_PERNUM").Value), "-", "");

//            // UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = sRecordset.Fields.Item(0).Value;
//            // // 2010.03.02 최동권 수정
//            // // 자사정보만 있고, 소득정보가 없어서 C레코드 없이 B레코드만 생성되는 경우가 있으므로
//            // // 소득자료가 없으면 B레코드가 생성되지 않도록 함(Nefs)
//            switch (File_Create_BRecord)
//            {
//                case 0:
//                    {
//                        Matrix_AddRow(CLTCOD + "- 징수의무자의 집계 레코드 생성 완료!", ref true);
//                        // / C레코드: 근로 주(현)근무처 레코드 /***********************************************/
//                        NEWCNT = 0;
//                        OLDCNT = 0;
//                        if (File_Create_CRecord == false)
//                        {
//                            ErrNum = 4;
//                            goto Error_Message;
//                        }
//                        Matrix_AddRow(CLTCOD + "- 징수의무자의 데이터 레코드" + NEWCNT + "건 생성 완료!", ref true);
//                        break;
//                    }

//                case 1:
//                    {
//                        ErrNum = 3;
//                        goto Error_Message;
//                        break;
//                    }

//                case 2 // // 해당자사에 소득자료가 없으면 B,C,D,E레코드 생성을 건너뜀
//         :
//                    {
//                        break;
//                    }
//            }
//            // /
//            sRecordset.MoveNext();
//        }
//        FileSystem.FileClose(1);
//        oForm.DataSources.UserDataSources.Item("Path").Value = FILNAM.Value;
//        Sbo_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//        File_Create = true;
//        // UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        sRecordset = null;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        sRecordset = null;
//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("A레코드(근로 제출자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("B레코드(근로 원천징수의무자별 집계 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 4)
//            Sbo_Application.StatusBar.SetText("C레코드(근로 주(현)근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 5)
//            Sbo_Application.StatusBar.SetText("CreateFolder Error : " + oStr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Sbo_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        File_Create = false;
//    }
//    private bool File_Create_ARecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 120991
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string PRTDAT;
//        string BUSNUM;
//        string CheckA;

//        CheckA = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        PRTDAT = Mid(oForm.Items.Item("PRTDAT").Specific.String, 1, 4) + Mid(oForm.Items.Item("PRTDAT").Specific.String, 6, 2) + Mid(oForm.Items.Item("PRTDAT").Specific.String, 9, 2);

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        // / 신고의무자수
//        // // 자사정보만 있고, 소득정보가 없어서 C레코드 없이 B레코드만 생성되는 경우가 있으므로
//        // // 소득자료가 없으면 B레코드가 생성되지 않도록 함
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 " + "WHERE U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "' " + "AND Code IN (SELECT U_CLTCOD FROM [@ZPY504H] WHERE U_JSNYER = '" + oJsnYear + "')";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount > 0)
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            BUSTOT = oRecordSet.Fields.Item(0).Value;
//        if (Conversion.Val(System.Convert.ToString(BUSTOT)) == 0)
//        {
//            ErrNum = 3;
//            goto Error_Message;
//        }
//        // / 업체정보
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "'";

//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }
//        else
//        {
//            // / 파일명
//            BUSNUM = Replace(oRecordSet.Fields.Item("U_BUSNUM").Value, "-", "");
//            if (Strings.Len(Strings.Trim(BUSNUM)) != 10)
//            {
//                ErrNum = 2;
//                goto Error_Message;
//            }
//            FILNAM.Value = oFilePath + "C" + Strings.Mid(BUSNUM, 1, 7) + "." + Strings.Mid(BUSNUM, 8, 3);
//            // / A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            Arec.RECGBN = "A"; // 
//            Arec.DTAGBN = "20";
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Arec.TAXCOD = oRecordSet.Fields.Item("U_TAXCODE").Value;
//            Arec.PRTDAT = PRTDAT;
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Arec.RPTGBN = oRecordSet.Fields.Item("U_TaxDGbn").Value;
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Arec.TAXAGE = oRecordSet.Fields.Item("U_TaxDCode").Value;
//            Arec.HOMTID = Trim(oRecordSet.Fields.Item("U_HOMETID").Value);
//            Arec.PGMCOD = "9000";
//            Arec.BUSNBR = Replace(oRecordSet.Fields.Item("U_TAXDBUS").Value, "-", "");
//            Arec.SANGHO = Trim(oRecordSet.Fields.Item("U_TAXDNAM").Value);
//            Arec.DAMDPT = Trim(oRecordSet.Fields.Item("U_CHGDPT").Value);
//            Arec.DAMNAM = Trim(oRecordSet.Fields.Item("U_CHGNAME").Value);
//            Arec.DAMTEL = Trim(oRecordSet.Fields.Item("U_CHGTEL").Value);
//            Arec.BUSCNT = VB6.Format(BUSTOT, new string("0", Strings.Len(Arec.BUSCNT))); // / 원천징수의무자수
//            Arec.HANCOD = "101";
//            Arec.FILLER = Strings.Space(Strings.Len(Arec.FILLER));

//            // / 필수입력 체크
//            if (Strings.Trim(Arec.TAXCOD) == "")
//            {
//                Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.RPTGBN) == "")
//            {
//                Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.BUSNBR) == "")
//            {
//                Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.SANGHO) == "")
//            {
//                Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.DAMDPT) == "")
//            {
//                Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.DAMNAM) == "")
//            {
//                Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Arec.DAMTEL) == "")
//            {
//                Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//            if (System.Convert.ToDouble(Arec.BUSCNT) == 0)
//            {
//                Matrix_AddRow("A레코드:신고내역이 존재하는 B레코드가 없습니다. 확인하여 주십시오.", ref true, ref true);
//                CheckA = System.Convert.ToString(true);
//            }
//        }

//        if (System.Convert.ToBoolean(CheckA) == false)
//            File_Create_ARecord = true;
//        else
//            File_Create_ARecord = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("귀속년도의 자사정보가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("자사정보등록의 사업자번호가 올바르지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("자사정보등록의 신고의무사업장이 존재하지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("A레코드오류: " + Information.Err.Description, ref false, ref true);
//        File_Create_ARecord = false;
//    }

//    // //------------------------------------------------------------
//    // // 반환값 리스트
//    // // 0 : 정상적으로 레코드 생성
//    // // 1 : 에러
//    // // 2 : B ~ E 레코드 생성 안함
//    // //------------------------------------------------------------
//    private short File_Create_BRecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 127517
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string CheckB;
//        CheckB = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / 집계정보
//        sQry = "EXEC ZPY521 'B', " + "'" + oJsnYear + "', '" + JSNGBN + "', '" + STRMON + "', '" + ENDMON + "', " + "'" + CLTCOD + "', '" + MSTBRK + "', '" + DPTSTR + "', '" + DPTEND + "','" + MSTCOD + "'";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }
//        else if (oRecordSet.Fields.Item("SUM_NEWCNT").Value == 0)
//        {
//            ErrNum = 2;
//            goto Error_Message;
//        }
//        else
//        {
//            // / B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            BUSCNT = BUSCNT + 1;

//            Brec.RECGBN = "B";
//            Brec.DTAGBN = "20";
//            Brec.BUSCNT = VB6.Format(BUSCNT, new string("0", Strings.Len(Brec.BUSCNT))); // / 원천징수의무자수 일련번호

//            Brec.NEWCNT = VB6.Format(oRecordSet.Fields.Item("SUM_NEWCNT").Value, new string("0", Strings.Len(Brec.NEWCNT))); // C Record수
//            Brec.OLDCNT = VB6.Format(oRecordSet.Fields.Item("SUM_OLDCNT").Value, new string("0", Strings.Len(Brec.OLDCNT))); // D Record수
//            Brec.INCOME = VB6.Format(oRecordSet.Fields.Item("SUM_INCOME").Value, new string("0", Strings.Len(Brec.INCOME)));
//            Brec.GULGAB = VB6.Format(oRecordSet.Fields.Item("SUM_GULGAB").Value, new string("0", Strings.Len(Brec.GULGAB)));
//            Brec.GULCOM = VB6.Format(0, new string("0", Strings.Len(Brec.GULCOM)));
//            Brec.GULJUM = VB6.Format(oRecordSet.Fields.Item("SUM_GULJUM").Value, new string("0", Strings.Len(Brec.GULJUM)));
//            Brec.GULNON = VB6.Format(oRecordSet.Fields.Item("SUM_GULNON").Value, new string("0", Strings.Len(Brec.GULNON)));
//            Brec.GULTOT = VB6.Format(oRecordSet.Fields.Item("SUM_GULGAB").Value + oRecordSet.Fields.Item("SUM_GULJUM").Value + oRecordSet.Fields.Item("SUM_GULNON").Value, new string("0", Strings.Len(Brec.GULTOT)));
//            // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Brec.RNGCOD = oForm.Items.Item("JSNTYP").Specific.Selected.Value;

//            Brec.FILLER = Strings.Space(Strings.Len(Brec.FILLER));

//            PrintLine(1, MDC_SetMod.sStr(Brec.RECGBN) + MDC_SetMod.sStr(Brec.DTAGBN) + MDC_SetMod.sStr(Brec.TAXCOD) + MDC_SetMod.sStr(Brec.BUSCNT) + MDC_SetMod.sStr(Brec.BUSNBR) + MDC_SetMod.sStr(Brec.SANGHO) + MDC_SetMod.sStr(Brec.COMPRT) + MDC_SetMod.sStr(Brec.PERNBR) + MDC_SetMod.sStr(Brec.NEWCNT) + MDC_SetMod.sStr(Brec.OLDCNT) + MDC_SetMod.sStr(Brec.INCOME) + MDC_SetMod.sStr(Brec.GULGAB) + MDC_SetMod.sStr(Brec.GULCOM) + MDC_SetMod.sStr(Brec.GULJUM) + MDC_SetMod.sStr(Brec.GULNON) + MDC_SetMod.sStr(Brec.GULTOT) + MDC_SetMod.sStr(Brec.RNGCOD) + MDC_SetMod.sStr(Brec.FILLER));

//            // / 필수입력 체크
//            if (Strings.Trim(Brec.BUSNBR) == "")
//            {
//                Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckB = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Brec.COMPRT) == "")
//            {
//                Matrix_AddRow("B레코드:대표자명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckB = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Brec.PERNBR) == "")
//            {
//                Matrix_AddRow("B레코드:법인(주민)등록번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckB = System.Convert.ToString(true);
//            }
//        }

//        if (System.Convert.ToBoolean(CheckB) == false)
//            File_Create_BRecord = 0;
//        else
//            File_Create_BRecord = 1;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//        {
//            Sbo_Application.StatusBar.SetText("집계레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            File_Create_BRecord = 1;
//        }
//        else if (ErrNum == 2)
//            File_Create_BRecord = 2;
//        else
//        {
//            Matrix_AddRow("B레코드오류: " + Information.Err.Description, ref false);
//            File_Create_BRecord = 1;
//        }
//    }
//    private bool File_Create_CRecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 131798
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string CheckC;
//        double OLDBIG;
//        double PILTOT;

//        CheckC = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        C_MSTCOD = "";
//        C_CLTCOD = "";
//        C_BUYCNT = System.Convert.ToString(0);

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / 사원 정보 조회

//        sQry = "EXEC ZPY521 'C', '" + oJsnYear + "', '" + JSNGBN + "', '" + STRMON + "', '" + ENDMON + "', " + "'" + CLTCOD + "', '" + MSTBRK + "', '" + DPTSTR + "', '" + DPTEND + "','" + MSTCOD + "'";

//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//        }
//        while (!oRecordSet.EOF)
//        {
//            NEWCNT = NEWCNT + 1;
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_MSTCOD = oRecordSet.Fields.Item("U_MSTCOD").Value;
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_CLTCOD = oRecordSet.Fields.Item("U_CLTCOD").Value;
//            // / 근로소득 주(현) 근무처 레코드 /
//            // / C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            Crec.RECGBN = "C";
//            Crec.DTAGBN = "20";
//            Crec.TAXCOD = Brec.TAXCOD;
//            Crec.SQNNBR = VB6.Format(NEWCNT, new string("0", Strings.Len(Crec.SQNNBR))); // / 일련번호
//            Crec.BUSNBR = Brec.BUSNBR;
//            Crec.JONCNT = VB6.Format(oRecordSet.Fields.Item("SUM_OLDCNT").Value, new string("0", Strings.Len(Crec.JONCNT))); // / 종전근무처수
//            Crec.DWEGBN = VB6.Format(oRecordSet.Fields.Item("U_DWEGBN").Value, new string("0", Strings.Len(Crec.DWEGBN)));
//            Crec.PERNBR = Replace(oRecordSet.Fields.Item("U_PERNBR").Value, "-", "");
//            Crec.INTGBN = VB6.Format(oRecordSet.Fields.Item("U_INTGBN").Value, new string("0", Strings.Len(Crec.INTGBN)));
//            if (Crec.DWEGBN == "1")
//                Crec.RGNCOD = "";
//            else
//                // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                Crec.RGNCOD = oRecordSet.Fields.Item("U_DWECOD").Value;// / 거주지국
//            if (Crec.INTGBN == "1")
//                Crec.GUKCOD = "";
//            else
//                // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                Crec.GUKCOD = oRecordSet.Fields.Item("U_GUKCOD").Value;// / 국적코드
//                                                                       // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Crec.HUSMAN = oRecordSet.Fields.Item("U_HUSMAN").Value; // / 세대주여부
//                                                                    // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Crec.JSNGBN = oRecordSet.Fields.Item("U_JSNGBN").Value; // / 정산구분
//                                                                    // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Crec.RGNTAX = oRecordSet.Fields.Item("U_FRGTAX").Value; // / 외국인단일세율
//            Crec.BUSNB1 = Replace(oRecordSet.Fields.Item("U_BUSNUM").Value, "-", ""); // / 주(현)사업자번호
//                                                                                      // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Crec.SANGHO = oRecordSet.Fields.Item("U_CLTNAM").Value; // / 주(현) 근무처명

//            Crec.STRINT = VB6.Format(Mid(Replace(oRecordSet.Fields.Item("U_STRINT").Value, "-", ""), 1, 8), new string("0", Strings.Len(Crec.STRINT)));
//            Crec.ENDINT = VB6.Format(Mid(Replace(oRecordSet.Fields.Item("U_ENDINT").Value, "-", ""), 1, 8), new string("0", Strings.Len(Crec.ENDINT)));
//            // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            Crec.MSTNAM = oRecordSet.Fields.Item("U_MSTNAM").Value;
//            Crec.STRGAM = VB6.Format(0, new string("0", Strings.Len(Crec.STRGAM)));
//            Crec.ENDGAM = VB6.Format(0, new string("0", Strings.Len(Crec.ENDGAM)));

//            Crec.PAYAMT = VB6.Format(oRecordSet.Fields.Item("U_PAYAMT").Value, new string("0", Strings.Len(Crec.PAYAMT)));
//            Crec.BNSAMT = VB6.Format(oRecordSet.Fields.Item("U_BNSAMT").Value, new string("0", Strings.Len(Crec.BNSAMT)));
//            Crec.INJBNS = VB6.Format(oRecordSet.Fields.Item("U_INBAMT").Value, new string("0", Strings.Len(Crec.INJBNS)));
//            Crec.JUSBNS = VB6.Format(oRecordSet.Fields.Item("U_JUSAMT").Value, new string("0", Strings.Len(Crec.JUSBNS))); // /2007
//            Crec.URIBNS = VB6.Format(oRecordSet.Fields.Item("U_URIAMT").Value, new string("0", Strings.Len(Crec.URIBNS))); // /2009
//            Crec.FILD01 = VB6.Format(0, new string("0", Strings.Len(Crec.FILD01)));
//            Crec.TOTAMT = VB6.Format(oRecordSet.Fields.Item("U_TOTAMT").Value, new string("0", Strings.Len(Crec.TOTAMT)));

//            OLDBIG = Val(oRecordSet.Fields.Item("U_BIGWA1").Value) + Val(oRecordSet.Fields.Item("U_BIGWA3").Value) + Val(oRecordSet.Fields.Item("U_BIGWA5").Value) + Val(oRecordSet.Fields.Item("U_BIGWA6").Value) + Val(oRecordSet.Fields.Item("U_BIGWU3").Value);

//            Crec.BIGG01 = VB6.Format(oRecordSet.Fields.Item("U_BTXG01").Value, new string("0", Strings.Len(Crec.BIGG01)));
//            Crec.BIGH01 = VB6.Format(oRecordSet.Fields.Item("U_BTXH01").Value, new string("0", Strings.Len(Crec.BIGH01)));
//            Crec.BIGH05 = VB6.Format(oRecordSet.Fields.Item("U_BTXH05").Value, new string("0", Strings.Len(Crec.BIGH05)));
//            Crec.BIGH06 = VB6.Format(oRecordSet.Fields.Item("U_BTXH06").Value, new string("0", Strings.Len(Crec.BIGH06)));
//            Crec.BIGH07 = VB6.Format(oRecordSet.Fields.Item("U_BTXH07").Value, new string("0", Strings.Len(Crec.BIGH07)));
//            Crec.BIGH08 = VB6.Format(oRecordSet.Fields.Item("U_BTXH08").Value, new string("0", Strings.Len(Crec.BIGH08)));
//            Crec.BIGH09 = VB6.Format(oRecordSet.Fields.Item("U_BTXH09").Value, new string("0", Strings.Len(Crec.BIGH09)));
//            Crec.BIGH10 = VB6.Format(oRecordSet.Fields.Item("U_BTXH10").Value, new string("0", Strings.Len(Crec.BIGH10)));
//            Crec.BIGH11 = VB6.Format(oRecordSet.Fields.Item("U_BTXH11").Value, new string("0", Strings.Len(Crec.BIGH11)));
//            Crec.BIGH12 = VB6.Format(oRecordSet.Fields.Item("U_BTXH12").Value, new string("0", Strings.Len(Crec.BIGH12)));
//            Crec.BIGH13 = VB6.Format(oRecordSet.Fields.Item("U_BTXH13").Value, new string("0", Strings.Len(Crec.BIGH13)));
//            Crec.BIGI01 = VB6.Format(oRecordSet.Fields.Item("U_BTXI01").Value, new string("0", Strings.Len(Crec.BIGI01)));
//            Crec.BIGK01 = VB6.Format(oRecordSet.Fields.Item("U_BTXK01").Value, new string("0", Strings.Len(Crec.BIGK01)));
//            Crec.BIGM01 = VB6.Format(oRecordSet.Fields.Item("U_BTXM01").Value, new string("0", Strings.Len(Crec.BIGM01)));
//            Crec.BIGM02 = VB6.Format(oRecordSet.Fields.Item("U_BTXM02").Value, new string("0", Strings.Len(Crec.BIGM02)));
//            Crec.BIGM03 = VB6.Format(oRecordSet.Fields.Item("U_BTXM03").Value, new string("0", Strings.Len(Crec.BIGM03)));
//            Crec.BIGO01 = VB6.Format(oRecordSet.Fields.Item("U_BTXO01").Value, new string("0", Strings.Len(Crec.BIGO01)));
//            Crec.BIGQ01 = VB6.Format(oRecordSet.Fields.Item("U_BTXQ01").Value, new string("0", Strings.Len(Crec.BIGQ01)));
//            Crec.BIGR10 = VB6.Format(oRecordSet.Fields.Item("U_BTXY01").Value, new string("0", Strings.Len(Crec.BIGR10))); // / 2011 Y01->R10으로 안쓰는필드대체
//            Crec.BIGS01 = VB6.Format(oRecordSet.Fields.Item("U_BTXS01").Value, new string("0", Strings.Len(Crec.BIGS01)));
//            Crec.BIGT01 = VB6.Format(oRecordSet.Fields.Item("U_BTXT01").Value, new string("0", Strings.Len(Crec.BIGT01)));
//            // Crec.BIGX01 = Format$(oRecordSet.Fields("U_BTXX01").Value, String$(Len(Crec.BIGX01), "0"))
//            Crec.BIGY02 = VB6.Format(oRecordSet.Fields.Item("U_BTXY02").Value, new string("0", Strings.Len(Crec.BIGY02)));
//            Crec.BIGY03 = VB6.Format(oRecordSet.Fields.Item("U_BTXY03").Value, new string("0", Strings.Len(Crec.BIGY03)));
//            Crec.BIGY21 = VB6.Format(oRecordSet.Fields.Item("U_BTXY21").Value, new string("0", Strings.Len(Crec.BIGY21)));
//            Crec.BIGZ01 = VB6.Format(oRecordSet.Fields.Item("U_BTXZ01").Value, new string("0", Strings.Len(Crec.BIGZ01)));
//            Crec.BIGY22 = VB6.Format(oRecordSet.Fields.Item("U_BTXY20").Value, new string("0", Strings.Len(Crec.BIGY22))); // / 그밖의비과
//            Crec.BIGTOT = VB6.Format(oRecordSet.Fields.Item("U_BTXTOT").Value, new string("0", Strings.Len(Crec.BIGTOT)));
//            Crec.BIGTO1 = VB6.Format(oRecordSet.Fields.Item("U_BTXTO1").Value, new string("0", Strings.Len(Crec.BIGTO1)));

//            Crec.INCOME = VB6.Format(oRecordSet.Fields.Item("U_INCOME").Value, new string("0", Strings.Len(Crec.INCOME)));
//            Crec.PILGNL = VB6.Format(oRecordSet.Fields.Item("U_PILGNL").Value, new string("0", Strings.Len(Crec.PILGNL)));
//            Crec.GNLOSD = VB6.Format(oRecordSet.Fields.Item("U_GNLOSD").Value, new string("0", Strings.Len(Crec.GNLOSD)));
//            Crec.INJBAS = VB6.Format(oRecordSet.Fields.Item("U_INJBAS").Value, new string("0", Strings.Len(Crec.INJBAS)));
//            Crec.INJBWO = VB6.Format(oRecordSet.Fields.Item("U_INJBWO").Value, new string("0", Strings.Len(Crec.INJBWO)));
//            Crec.BUYNSU = VB6.Format(oRecordSet.Fields.Item("U_BUYNSU").Value, new string("0", Strings.Len(Crec.BUYNSU)));
//            Crec.INJBYN = VB6.Format(oRecordSet.Fields.Item("U_INJBYN").Value, new string("0", Strings.Len(Crec.INJBYN)));
//            Crec.GYNGLO = VB6.Format(oRecordSet.Fields.Item("U_GYNGLO").Value, new string("0", Strings.Len(Crec.GYNGLO)));
//            Crec.INJGYN = VB6.Format(oRecordSet.Fields.Item("U_INJGYN").Value, new string("0", Strings.Len(Crec.INJGYN)));
//            Crec.JANGAE = VB6.Format(oRecordSet.Fields.Item("U_JANGAE").Value, new string("0", Strings.Len(Crec.JANGAE)));
//            Crec.INJJAE = VB6.Format(oRecordSet.Fields.Item("U_INJJAE").Value, new string("0", Strings.Len(Crec.INJJAE)));
//            Crec.INJBNY = VB6.Format(oRecordSet.Fields.Item("U_INJBNJ").Value, new string("0", Strings.Len(Crec.INJBNY)));
//            Crec.BUYN06 = VB6.Format(oRecordSet.Fields.Item("U_BUYN06").Value, new string("0", Strings.Len(Crec.BUYN06)));
//            Crec.INJSON = VB6.Format(oRecordSet.Fields.Item("U_INJSON").Value, new string("0", Strings.Len(Crec.INJSON)));
//            Crec.CHLSAN = VB6.Format(oRecordSet.Fields.Item("U_CHLSAN").Value, new string("0", Strings.Len(Crec.CHLSAN)));
//            Crec.INJCHL = VB6.Format(oRecordSet.Fields.Item("U_INJCHL").Value, new string("0", Strings.Len(Crec.INJCHL)));
//            Crec.FILD02 = VB6.Format(0, new string("0", Strings.Len(Crec.FILD02)));
//            Crec.DAGYSU = VB6.Format(oRecordSet.Fields.Item("U_DAGYSU").Value, new string("0", Strings.Len(Crec.DAGYSU))); // / 다자녀수(2007)
//            Crec.INJADD = VB6.Format(oRecordSet.Fields.Item("U_INJADD").Value, new string("0", Strings.Len(Crec.INJADD))); // / 다자녀공제금액
//            Crec.KUKGON = VB6.Format(oRecordSet.Fields.Item("U_KUKGON").Value, new string("0", Strings.Len(Crec.KUKGON))); // / 국민연금보험료공제
//            Crec.YUNGON = VB6.Format(oRecordSet.Fields.Item("U_YUNGON").Value, new string("0", Strings.Len(Crec.YUNGON))); // / 연금보험료공제(공무원)
//            Crec.YUNGO1 = VB6.Format(oRecordSet.Fields.Item("U_YUNGO1").Value, new string("0", Strings.Len(Crec.YUNGO1))); // /               (군인)
//            Crec.YUNGO2 = VB6.Format(oRecordSet.Fields.Item("U_YUNGO2").Value, new string("0", Strings.Len(Crec.YUNGO2))); // /               (사립)
//            Crec.YUNGO3 = VB6.Format(oRecordSet.Fields.Item("U_YUNGO3").Value, new string("0", Strings.Len(Crec.YUNGO3))); // /               (별정우체국)
//            Crec.GITRE2 = VB6.Format(oRecordSet.Fields.Item("U_GITRE2").Value, new string("0", Strings.Len(Crec.GITRE2))); // / 퇴직연금(과학기술인)
//            Crec.GITRET = VB6.Format(oRecordSet.Fields.Item("U_GITRET").Value, new string("0", Strings.Len(Crec.GITRET))); // /         (근로자퇴직급여보장법)

//            Crec.PILMBH = VB6.Format(oRecordSet.Fields.Item("U_PILMBH").Value, new string("0", Strings.Len(Crec.PILMBH)));
//            Crec.PILGBH = VB6.Format(oRecordSet.Fields.Item("U_PILGBH").Value, new string("0", Strings.Len(Crec.PILGBH)));
//            Crec.PILBHM = VB6.Format(oRecordSet.Fields.Item("U_PILBHM").Value, new string("0", Strings.Len(Crec.PILBHM)));
//            Crec.PILJHM = VB6.Format(oRecordSet.Fields.Item("U_PILJHM").Value, new string("0", Strings.Len(Crec.PILJHM)));
//            Crec.PILMED = VB6.Format(oRecordSet.Fields.Item("U_PILMED").Value, new string("0", Strings.Len(Crec.PILMED)));
//            Crec.PILSCH = VB6.Format(oRecordSet.Fields.Item("U_PILSCH").Value, new string("0", Strings.Len(Crec.PILSCH)));
//            Crec.PILHUS = VB6.Format(oRecordSet.Fields.Item("U_PILHUS").Value, new string("0", Strings.Len(Crec.PILHUS))); // / 주택임대차차입금원리상환공제금액
//            Crec.PILHU2 = VB6.Format(oRecordSet.Fields.Item("U_PILHU2").Value, new string("0", Strings.Len(Crec.PILHU2))); // / 주택임대차차입금원리상환공제금액-거주자(분리) 2011
//            Crec.PILWOL = VB6.Format(oRecordSet.Fields.Item("U_PILWOL").Value, new string("0", Strings.Len(Crec.PILWOL))); // / 월세액
//            Crec.PILJHE = VB6.Format(oRecordSet.Fields.Item("U_PILJHE").Value, new string("0", Strings.Len(Crec.PILJHE))); // / 장기주택저당차입금이자상환공제(2008)
//            Crec.PILJH2 = VB6.Format(oRecordSet.Fields.Item("U_PILJH2").Value, new string("0", Strings.Len(Crec.PILJH2))); // / 장기주택저당차입금이자상환공제(2011 분리)
//            Crec.PILJH3 = VB6.Format(oRecordSet.Fields.Item("U_PILJH3").Value, new string("0", Strings.Len(Crec.PILJH3))); // / 장기주택저당차입금이자상환공제(2011 분리)
//            Crec.PILGBU = VB6.Format(oRecordSet.Fields.Item("U_PILGBU").Value, new string("0", Strings.Len(Crec.PILGBU)));
//            // Crec.PILHUN = Format$(oRecordSet.Fields("U_PILHUN").Value, String$(Len(Crec.PILHUN), "0"))  '/ 2009년 혼인장례 제거
//            Crec.PILFLD = VB6.Format(0, new string("0", Strings.Len(Crec.PILFLD)));
//            if (oRecordSet.Fields.Item("U_PILTOT").Value > 0)
//                Crec.PILTOT = VB6.Format(oRecordSet.Fields.Item("U_PILTOT").Value, new string("0", Strings.Len(Crec.PILTOT)));
//            else
//            {
//                PILTOT = Val(oRecordSet.Fields.Item("U_PILMBH").Value) + Val(oRecordSet.Fields.Item("U_PILGBH").Value) + Val(oRecordSet.Fields.Item("U_PILBHM").Value) + Val(oRecordSet.Fields.Item("U_PILJHM").Value) + Val(oRecordSet.Fields.Item("U_PILMED").Value) + Val(oRecordSet.Fields.Item("U_PILSCH").Value) + Val(oRecordSet.Fields.Item("U_PILHUS").Value) + Val(oRecordSet.Fields.Item("U_PILHU2").Value) + Val(oRecordSet.Fields.Item("U_PILWOL").Value) + Val(oRecordSet.Fields.Item("U_PILJHE").Value) + Val(oRecordSet.Fields.Item("U_PILJH2").Value) + Val(oRecordSet.Fields.Item("U_PILJH3").Value) + Val(oRecordSet.Fields.Item("U_PILGBU").Value);
//                Crec.PILTOT = VB6.Format(PILTOT, new string("0", Strings.Len(Crec.PILTOT)));
//            }
//            Crec.PILGON = VB6.Format(oRecordSet.Fields.Item("U_PILGON").Value, new string("0", Strings.Len(Crec.PILGON)));
//            Crec.CHAGAM = VB6.Format(oRecordSet.Fields.Item("U_CHAGAM").Value, new string("0", Strings.Len(Crec.CHAGAM)));
//            Crec.GITGYN = VB6.Format(oRecordSet.Fields.Item("U_GITGYN").Value, new string("0", Strings.Len(Crec.GITGYN)));
//            Crec.GITYUN = VB6.Format(oRecordSet.Fields.Item("U_GITYUN").Value, new string("0", Strings.Len(Crec.GITYUN)));
//            Crec.GITSGI = VB6.Format(oRecordSet.Fields.Item("U_GITSGI").Value, new string("0", Strings.Len(Crec.GITSGI)));
//            Crec.GITHUS = VB6.Format(oRecordSet.Fields.Item("U_GITHUS").Value, new string("0", Strings.Len(Crec.GITHUS))); // /주택마련저축(청약)
//            Crec.GITHU1 = VB6.Format(oRecordSet.Fields.Item("U_GITHU1").Value, new string("0", Strings.Len(Crec.GITHU1))); // /            (종합저축)
//            Crec.GITHU2 = VB6.Format(oRecordSet.Fields.Item("U_GITHU2").Value, new string("0", Strings.Len(Crec.GITHU2))); // /            (장기주택)
//            Crec.GITHU3 = VB6.Format(oRecordSet.Fields.Item("U_GITHU3").Value, new string("0", Strings.Len(Crec.GITHU3))); // /            (근로자)
//            Crec.GITINV = VB6.Format(oRecordSet.Fields.Item("U_GITINV").Value, new string("0", Strings.Len(Crec.GITINV)));
//            Crec.GITCAD = VB6.Format(oRecordSet.Fields.Item("U_GITCAD").Value, new string("0", Strings.Len(Crec.GITCAD)));
//            // If Val(oRecordSet.Fields("U_GITUSJ").Value) < 0 Then
//            // Crec.GITUSG = 1 '우리사주조합소득공제(음수1 양수0)
//            // Else
//            // Crec.GITUSG = 0
//            // End If
//            Crec.GITUSJ = VB6.Format(oRecordSet.Fields.Item("U_GITUSJ").Value, new string("0", Strings.Len(Crec.GITUSJ)));
//            Crec.GITJFD = VB6.Format(oRecordSet.Fields.Item("U_GITJFD").Value, new string("0", Strings.Len(Crec.GITJFD)));
//            Crec.GITGYU = VB6.Format(oRecordSet.Fields.Item("U_GITGYU").Value, new string("0", Strings.Len(Crec.GITJFD)));
//            Crec.FILD03 = VB6.Format(0, new string("0", Strings.Len(Crec.FILD03)));
//            // If Val(oRecordSet.Fields("U_GITTOT").Value) < 0 Then
//            // Crec.GITTOG = 1
//            // Else
//            // Crec.GITTOG = 0 '기타소득공제계(기호 음수1, 양수0)
//            // End If
//            Crec.GITTOT = VB6.Format(oRecordSet.Fields.Item("U_GITTOT").Value, new string("0", Strings.Len(Crec.GITTOT)));
//            Crec.TAXSTD = VB6.Format(oRecordSet.Fields.Item("U_TAXSTD").Value, new string("0", Strings.Len(Crec.TAXSTD)));
//            Crec.SANTAX = VB6.Format(oRecordSet.Fields.Item("U_SANTAX").Value, new string("0", Strings.Len(Crec.SANTAX)));
//            Crec.GAMSOD = VB6.Format(oRecordSet.Fields.Item("U_GAMSOD").Value, new string("0", Strings.Len(Crec.GAMSOD)));
//            Crec.GAMJOS = VB6.Format(oRecordSet.Fields.Item("U_GAMJOS").Value, new string("0", Strings.Len(Crec.GAMJOS)));
//            Crec.GAMJYK = VB6.Format(oRecordSet.Fields.Item("U_GAMJYK").Value, new string("0", Strings.Len(Crec.GAMJYK)));
//            Crec.GAMFLD = new string("0", Strings.Len(Crec.GAMFLD));
//            Crec.GAMTOT = VB6.Format(oRecordSet.Fields.Item("U_GAMTOT").Value, new string("0", Strings.Len(Crec.GAMTOT)));
//            Crec.TAXGNL = VB6.Format(oRecordSet.Fields.Item("U_TAXGNL").Value, new string("0", Strings.Len(Crec.TAXGNL)));
//            Crec.NABSEE = VB6.Format(oRecordSet.Fields.Item("U_TAXNAB").Value, new string("0", Strings.Len(Crec.NABSEE)));
//            Crec.TAXBRO = VB6.Format(oRecordSet.Fields.Item("U_TAXBRO").Value, new string("0", Strings.Len(Crec.TAXBRO)));
//            Crec.TAXGBU = VB6.Format(oRecordSet.Fields.Item("U_TAXGBU").Value, new string("0", Strings.Len(Crec.TAXGBU)));
//            Crec.TAXFRG = VB6.Format(oRecordSet.Fields.Item("U_TAXFRG").Value, new string("0", Strings.Len(Crec.TAXFRG)));
//            Crec.FILD04 = VB6.Format(0, new string("0", Strings.Len(Crec.FILD04)));
//            Crec.TAXTOT = VB6.Format(oRecordSet.Fields.Item("U_TAXTOT").Value, new string("0", Strings.Len(Crec.TAXTOT)));
//            Crec.GULGAB = VB6.Format(oRecordSet.Fields.Item("U_GULGAB").Value, new string("0", Strings.Len(Crec.GULGAB)));
//            Crec.GULJUM = VB6.Format(oRecordSet.Fields.Item("U_GULJUM").Value, new string("0", Strings.Len(Crec.GULJUM)));
//            Crec.GULNON = VB6.Format(oRecordSet.Fields.Item("U_GULNON").Value, new string("0", Strings.Len(Crec.GULNON)));
//            // Crec.GULTOT = Format$(oRecordSet.Fields("U_GULGAB").Value + oRecordSet.Fields("U_GULJUM").Value + oRecordSet.Fields("U_GULNON").Value, String$(Len(Crec.GULTOT), "0"))
//            Crec.NANGAB = VB6.Format(oRecordSet.Fields.Item("U_NANGAB").Value, new string("0", Strings.Len(Crec.NANGAB)));
//            Crec.NANJUM = VB6.Format(oRecordSet.Fields.Item("U_NANJUM").Value, new string("0", Strings.Len(Crec.NANJUM)));
//            Crec.NANNON = VB6.Format(oRecordSet.Fields.Item("U_NANNON").Value, new string("0", Strings.Len(Crec.NANNON)));
//            // Crec.NANTOT = Format$(oRecordSet.Fields("U_NANGAB").Value + oRecordSet.Fields("U_NANJUM").Value + oRecordSet.Fields("U_NANNON").Value, String$(Len(Crec.NANTOT), "0"))
//            // / 차감징수세액(2009년추가)
//            if (Val(oRecordSet.Fields.Item("U_CHAGAB").Value) < 0)
//            {
//                Crec.CHAGAG = (string)1;
//                Crec.CHAGAB = VB6.Format(-1 * oRecordSet.Fields.Item("U_CHAGAB").Value, new string("0", Strings.Len(Crec.CHAGAB))); // 기현)차감소득세
//            }
//            else
//            {
//                Crec.CHAGAG = (string)0; // 현)차감소득세(기호 음수1, 양수0)
//                Crec.CHAGAB = VB6.Format(oRecordSet.Fields.Item("U_CHAGAB").Value, new string("0", Strings.Len(Crec.CHAGAB))); // 기현)차감소득세
//            }
//            if (Val(oRecordSet.Fields.Item("U_CHAJUM").Value) < 0)
//            {
//                Crec.CHAJUG = (string)1;
//                Crec.CHAJUM = VB6.Format(-1 * oRecordSet.Fields.Item("U_CHAJUM").Value, new string("0", Strings.Len(Crec.CHAJUM))); // 기현)차감소득세
//            }
//            else
//            {
//                Crec.CHAJUG = (string)0; // 현)차감주민세(기호 음수1, 양수0)
//                Crec.CHAJUM = VB6.Format(oRecordSet.Fields.Item("U_CHAJUM").Value, new string("0", Strings.Len(Crec.CHAJUM))); // 기현)차감소득세
//            }
//            if (Val(oRecordSet.Fields.Item("U_CHANON").Value) < 0)
//            {
//                Crec.CHANOG = (string)1;
//                Crec.CHANON = VB6.Format(-1 * oRecordSet.Fields.Item("U_CHANON").Value, new string("0", Strings.Len(Crec.CHANON))); // 기현)차감소득세
//            }
//            else
//            {
//                Crec.CHANOG = (string)0; // 현)차감주민세(기호 음수1, 양수0)
//                Crec.CHANON = VB6.Format(oRecordSet.Fields.Item("U_CHANON").Value, new string("0", Strings.Len(Crec.CHANON))); // 기현)차감소득세
//            }

//            if ((Val(oRecordSet.Fields.Item("U_CHAGAB").Value) + Val(oRecordSet.Fields.Item("U_CHAJUM").Value) + Val(oRecordSet.Fields.Item("U_CHANON").Value)) < 0)
//            {
//            }
//            else
//            {
//            }

//            Crec.FILLER = Strings.Space(Strings.Len(Crec.FILLER));


//            PrintLine(1, MDC_SetMod.sStr(Crec.RECGBN) + MDC_SetMod.sStr(Crec.DTAGBN) + MDC_SetMod.sStr(Crec.TAXCOD) + MDC_SetMod.sStr(Crec.SQNNBR) + MDC_SetMod.sStr(Crec.BUSNBR) + MDC_SetMod.sStr(Crec.JONCNT) + MDC_SetMod.sStr(Crec.DWEGBN) + MDC_SetMod.sStr(Crec.RGNCOD) + MDC_SetMod.sStr(Crec.RGNTAX) + MDC_SetMod.sStr(Crec.MSTNAM) + MDC_SetMod.sStr(Crec.INTGBN) + MDC_SetMod.sStr(Crec.PERNBR) + MDC_SetMod.sStr(Crec.GUKCOD) + MDC_SetMod.sStr(Crec.HUSMAN) + MDC_SetMod.sStr(Crec.JSNGBN) + MDC_SetMod.sStr(Crec.BUSNB1) + MDC_SetMod.sStr(Crec.SANGHO) + MDC_SetMod.sStr(Crec.STRINT) + MDC_SetMod.sStr(Crec.ENDINT) + MDC_SetMod.sStr(Crec.STRGAM) + MDC_SetMod.sStr(Crec.ENDGAM) + MDC_SetMod.sStr(Crec.PAYAMT) + MDC_SetMod.sStr(Crec.BNSAMT) + MDC_SetMod.sStr(Crec.INJBNS) + MDC_SetMod.sStr(Crec.JUSBNS) + MDC_SetMod.sStr(Crec.URIBNS) + MDC_SetMod.sStr(Crec.FILD01) + MDC_SetMod.sStr(Crec.TOTAMT) + MDC_SetMod.sStr(Crec.BIGG01) + MDC_SetMod.sStr(Crec.BIGH01) + MDC_SetMod.sStr(Crec.BIGH05) + MDC_SetMod.sStr(Crec.BIGH06) + MDC_SetMod.sStr(Crec.BIGH07) + MDC_SetMod.sStr(Crec.BIGH08) + MDC_SetMod.sStr(Crec.BIGH09) + MDC_SetMod.sStr(Crec.BIGH10) + MDC_SetMod.sStr(Crec.BIGH11) + MDC_SetMod.sStr(Crec.BIGH12) + MDC_SetMod.sStr(Crec.BIGH13) + MDC_SetMod.sStr(Crec.BIGI01) + MDC_SetMod.sStr(Crec.BIGK01) + MDC_SetMod.sStr(Crec.BIGM01) + MDC_SetMod.sStr(Crec.BIGM02) + MDC_SetMod.sStr(Crec.BIGM03) + MDC_SetMod.sStr(Crec.BIGO01) + MDC_SetMod.sStr(Crec.BIGQ01) + MDC_SetMod.sStr(Crec.BIGR10) + MDC_SetMod.sStr(Crec.BIGS01) + MDC_SetMod.sStr(Crec.BIGT01) + MDC_SetMod.sStr(Crec.BIGY02) + MDC_SetMod.sStr(Crec.BIGY03) + MDC_SetMod.sStr(Crec.BIGY21) + MDC_SetMod.sStr(Crec.BIGZ01) + MDC_SetMod.sStr(Crec.BIGY22) + MDC_SetMod.sStr(Crec.BIGTOT) + MDC_SetMod.sStr(Crec.BIGTO1) + MDC_SetMod.sStr(Crec.INCOME) + MDC_SetMod.sStr(Crec.PILGNL) + MDC_SetMod.sStr(Crec.GNLOSD) + MDC_SetMod.sStr(Crec.INJBAS) + MDC_SetMod.sStr(Crec.INJBWO) + MDC_SetMod.sStr(Crec.BUYNSU) + MDC_SetMod.sStr(Crec.INJBYN) + MDC_SetMod.sStr(Crec.GYNGLO) + MDC_SetMod.sStr(Crec.INJGYN) + MDC_SetMod.sStr(Crec.JANGAE) + MDC_SetMod.sStr(Crec.INJJAE) + MDC_SetMod.sStr(Crec.INJBNY) + MDC_SetMod.sStr(Crec.BUYN06) + MDC_SetMod.sStr(Crec.INJSON) + MDC_SetMod.sStr(Crec.CHLSAN) + MDC_SetMod.sStr(Crec.INJCHL) + MDC_SetMod.sStr(Crec.FILD02) + MDC_SetMod.sStr(Crec.DAGYSU) + MDC_SetMod.sStr(Crec.INJADD) + MDC_SetMod.sStr(Crec.KUKGON) + MDC_SetMod.sStr(Crec.YUNGON) + MDC_SetMod.sStr(Crec.YUNGO1) + MDC_SetMod.sStr(Crec.YUNGO2) + MDC_SetMod.sStr(Crec.YUNGO3) + MDC_SetMod.sStr(Crec.GITRE2) + MDC_SetMod.sStr(Crec.GITRET) + MDC_SetMod.sStr(Crec.PILMBH) + MDC_SetMod.sStr(Crec.PILGBH) + MDC_SetMod.sStr(Crec.PILBHM) + MDC_SetMod.sStr(Crec.PILJHM) + MDC_SetMod.sStr(Crec.PILMED) + MDC_SetMod.sStr(Crec.PILSCH) + MDC_SetMod.sStr(Crec.PILHUS) + MDC_SetMod.sStr(Crec.PILHU2) + MDC_SetMod.sStr(Crec.PILWOL) + MDC_SetMod.sStr(Crec.PILJHE) + MDC_SetMod.sStr(Crec.PILJH2) + MDC_SetMod.sStr(Crec.PILJH3) + MDC_SetMod.sStr(Crec.PILGBU) + MDC_SetMod.sStr(Crec.PILFLD) + MDC_SetMod.sStr(Crec.PILTOT) + MDC_SetMod.sStr(Crec.PILGON) + MDC_SetMod.sStr(Crec.CHAGAM) + MDC_SetMod.sStr(Crec.GITGYN) + MDC_SetMod.sStr(Crec.GITYUN) + MDC_SetMod.sStr(Crec.GITSGI) + MDC_SetMod.sStr(Crec.GITHUS) + MDC_SetMod.sStr(Crec.GITHU1) + MDC_SetMod.sStr(Crec.GITHU2) + MDC_SetMod.sStr(Crec.GITHU3) + MDC_SetMod.sStr(Crec.GITINV) + MDC_SetMod.sStr(Crec.GITCAD) + MDC_SetMod.sStr(Crec.GITUSJ) + MDC_SetMod.sStr(Crec.GITJFD) + MDC_SetMod.sStr(Crec.GITGYU) + MDC_SetMod.sStr(Crec.FILD03) + MDC_SetMod.sStr(Crec.GITTOT) + MDC_SetMod.sStr(Crec.TAXSTD) + MDC_SetMod.sStr(Crec.SANTAX) + MDC_SetMod.sStr(Crec.GAMSOD) + MDC_SetMod.sStr(Crec.GAMJOS) + MDC_SetMod.sStr(Crec.GAMJYK) + MDC_SetMod.sStr(Crec.GAMFLD) + MDC_SetMod.sStr(Crec.GAMTOT) + MDC_SetMod.sStr(Crec.TAXGNL) + MDC_SetMod.sStr(Crec.NABSEE) + MDC_SetMod.sStr(Crec.TAXBRO) + MDC_SetMod.sStr(Crec.TAXGBU) + MDC_SetMod.sStr(Crec.TAXFRG) + MDC_SetMod.sStr(Crec.FILD04) + MDC_SetMod.sStr(Crec.TAXTOT) + MDC_SetMod.sStr(Crec.GULGAB) + MDC_SetMod.sStr(Crec.GULJUM) + MDC_SetMod.sStr(Crec.GULNON) + MDC_SetMod.sStr(Crec.NANGAB) + MDC_SetMod.sStr(Crec.NANJUM) + MDC_SetMod.sStr(Crec.NANNON) + MDC_SetMod.sStr(Crec.CHAGAG) + MDC_SetMod.sStr(Crec.CHAGAB) + MDC_SetMod.sStr(Crec.CHAJUG) + MDC_SetMod.sStr(Crec.CHAJUM) + MDC_SetMod.sStr(Crec.CHANOG) + MDC_SetMod.sStr(Crec.CHANON) + MDC_SetMod.sStr(Crec.FILLER));



//            // / 기본공제수
//            if (Conversion.Val(Crec.INJBWO) > 0)
//                C_BUYCNT = System.Convert.ToString(1);
//            else
//                C_BUYCNT = System.Convert.ToString(0);
//            C_BUYCNT = System.Convert.ToString(System.Convert.ToDouble(C_BUYCNT) + Conversion.Val(Crec.BUYNSU));
//            Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 생성 완료.", ref true);
//            // / 필수입력 체크
//            if (Crec.DWEGBN == "0")
//            {
//                Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-거주자구분이 누락되었습니다. 확인하여 주십시오.", ref false, ref true);
//                CheckC = System.Convert.ToString(true);
//            }
//            if (Crec.DWEGBN != "1" & Strings.Trim(Crec.RGNCOD) == "")
//            {
//                Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-비거주자일경우 거주지국코드는 필수입니다. 확인하여 주십시오.", ref false, ref true);
//                CheckC = System.Convert.ToString(true);
//            }
//            if (Crec.INTGBN != "1" & Strings.Trim(Crec.GUKCOD) == "")
//            {
//                Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-외국인일경우 국적코드는 필수입니다. 확인하여 주십시오.", ref false, ref true);
//                CheckC = System.Convert.ToString(true);
//            }
//            if (Crec.INTGBN == "1" & Strings.Len(Strings.Trim(Crec.PERNBR)) != 13)
//            {
//                Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-주민등록번호를 확인하여 주십시오.", ref false, ref true);
//                CheckC = System.Convert.ToString(true);
//            }
//            if (OLDBIG != 0)
//            {
//                Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-비과세금액 세부코드이관작업을 하셔야합니다.", ref false, ref true);
//                CheckC = System.Convert.ToString(true);
//            }

//            // / D레코드: 종전근무처 레코드
//            if (Conversion.Val(Crec.JONCNT) > 0)
//            {
//                if (File_Create_DRecord == false)
//                {
//                    ErrNum = 2;
//                    goto Error_Message;
//                }
//            }
//            // / E레코드: 부양가족 레코드
//            // / 외국인은 단일세율을 적용할 경우 수록 안함
//            if (Crec.INTGBN == "1" | (Crec.INTGBN == "9" & Crec.RGNTAX == "2"))
//            {
//                if (File_Create_ERecord == false)
//                {
//                    ErrNum = 3;
//                    goto Error_Message;
//                }
//            }
//            if (Conversion.Val(Crec.GITRET) != 0 | Conversion.Val(Crec.GITRE2) != 0 | Conversion.Val(Crec.GITGYN) != 0 | Conversion.Val(Crec.GITYUN) != 0 | Conversion.Val(Crec.GITHUS) != 0 | Conversion.Val(Crec.GITHU1) != 0 | Conversion.Val(Crec.GITHU2) != 0 | Conversion.Val(Crec.GITHU3) != 0 | Conversion.Val(Crec.GITJFD) != 0)
//            {
//                if (File_Create_FRecord == false)
//                {
//                    ErrNum = 4;
//                    goto Error_Message;
//                }
//            }
//            oRecordSet.MoveNext();
//        }


//        if (System.Convert.ToBoolean(CheckC) == false)
//            File_Create_CRecord = true;
//        else
//            File_Create_CRecord = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("주(현)근무처 레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("D레코드(종전근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("E레코드(부양가족 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 4)
//            Sbo_Application.StatusBar.SetText("F레코드(연금.저축 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("C레코드오류: " + Information.Err.Description, ref false);
//        File_Create_CRecord = false;
//    }

//    private bool File_Create_DRecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 159913
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string CheckD;
//        short JONCNT;

//        CheckD = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / 종전근무지정보
//        sQry = "SELECT      T0.* ";
//        sQry = sQry + " FROM [@ZPY502L] T0 INNER JOIN [@ZPY502H] T1 ON T0.DocEntry = T1.DocEntry";
//        sQry = sQry + " WHERE    T1.U_JSNYER = '" + oJsnYear + "'";
//        sQry = sQry + " AND      T1.U_MSTCOD = '" + C_MSTCOD + "'";
//        sQry = sQry + " AND      T1.U_CLTCOD = '" + C_CLTCOD + "'";
//        sQry = sQry + " ORDER BY T0.LineID";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }
//        JONCNT = 0;
//        while (!oRecordSet.EOF)
//        {
//            // / D Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            OLDCNT = OLDCNT + 1; // / 전사원 종전근무지수
//            JONCNT = JONCNT + 1; // / 해당사원의 종전근무지일련번호
//            Drec.RECGBN = "D";
//            Drec.DTAGBN = "20";
//            Drec.TAXCOD = Arec.TAXCOD;
//            Drec.SQNNBR = VB6.Format(NEWCNT, new string("0", Strings.Len(Drec.SQNNBR)));
//            Drec.BUSNBR = Brec.BUSNBR;
//            // Drec.FILD01 = Format$(0, String$(Len(Drec.FILD01), "0"))
//            Drec.FILD01 = Strings.Space(Strings.Len(Drec.FILD01));
//            Drec.PERNBR = Crec.PERNBR;

//            Drec.TAXJOH = "2"; // // 납세조합 구분
//            Drec.JONNAM = Trim(oRecordSet.Fields.Item("U_JONNAM").Value);
//            Drec.JONNBR = Replace(oRecordSet.Fields.Item("U_JONNBR").Value, "-", "");
//            Drec.STRINT = Replace(oRecordSet.Fields.Item("U_JONSTR").Value, "-", "");
//            Drec.ENDINT = Replace(oRecordSet.Fields.Item("U_JONEND").Value, "-", "");
//            Drec.STRGAM = VB6.Format(Replace(oRecordSet.Fields.Item("U_JONGFR").Value, "-", ""), new string("0", Strings.Len(Drec.STRGAM)));
//            Drec.ENDGAM = VB6.Format(Replace(oRecordSet.Fields.Item("U_JONGTO").Value, "-", ""), new string("0", Strings.Len(Drec.STRGAM)));

//            Drec.PAYAMT = VB6.Format(oRecordSet.Fields.Item("U_JONPAY").Value, new string("0", Strings.Len(Drec.PAYAMT)));
//            Drec.BNSAMT = VB6.Format(oRecordSet.Fields.Item("U_JONBNS").Value, new string("0", Strings.Len(Drec.BNSAMT)));
//            Drec.INJBNS = VB6.Format(oRecordSet.Fields.Item("U_INJBNS").Value, new string("0", Strings.Len(Drec.INJBNS)));
//            Drec.JUSBNS = VB6.Format(oRecordSet.Fields.Item("U_JONJUS").Value, new string("0", Strings.Len(Drec.JUSBNS)));
//            Drec.URIBNS = VB6.Format(oRecordSet.Fields.Item("U_URIBNS").Value, new string("0", Strings.Len(Drec.URIBNS)));
//            Drec.FILD02 = VB6.Format(0, new string("0", Strings.Len(Drec.FILD02)));
//            Drec.TOTAMT = VB6.Format(oRecordSet.Fields.Item("U_JONPAY").Value + oRecordSet.Fields.Item("U_JONBNS").Value + oRecordSet.Fields.Item("U_INJBNS").Value + oRecordSet.Fields.Item("U_JONJUS").Value + oRecordSet.Fields.Item("U_URIBNS").Value, new string("0", Strings.Len(Drec.TOTAMT)));

//            Drec.BIGG01 = VB6.Format(oRecordSet.Fields.Item("U_JBTG01").Value, new string("0", Strings.Len(Drec.BIGG01)));
//            Drec.BIGH01 = VB6.Format(oRecordSet.Fields.Item("U_JBTH01").Value, new string("0", Strings.Len(Drec.BIGH01)));
//            Drec.BIGH05 = VB6.Format(oRecordSet.Fields.Item("U_JBTH05").Value, new string("0", Strings.Len(Drec.BIGH05)));
//            Drec.BIGH06 = VB6.Format(oRecordSet.Fields.Item("U_JBTH06").Value, new string("0", Strings.Len(Drec.BIGH06)));
//            Drec.BIGH07 = VB6.Format(oRecordSet.Fields.Item("U_JBTH07").Value, new string("0", Strings.Len(Drec.BIGH07)));
//            Drec.BIGH08 = VB6.Format(oRecordSet.Fields.Item("U_JBTH08").Value, new string("0", Strings.Len(Drec.BIGH08)));
//            Drec.BIGH09 = VB6.Format(oRecordSet.Fields.Item("U_JBTH09").Value, new string("0", Strings.Len(Drec.BIGH09)));
//            Drec.BIGH10 = VB6.Format(oRecordSet.Fields.Item("U_JBTH10").Value, new string("0", Strings.Len(Drec.BIGH10)));
//            Drec.BIGH11 = VB6.Format(oRecordSet.Fields.Item("U_JBTH11").Value, new string("0", Strings.Len(Drec.BIGH11)));
//            Drec.BIGH12 = VB6.Format(oRecordSet.Fields.Item("U_JBTH12").Value, new string("0", Strings.Len(Drec.BIGH12)));
//            Drec.BIGH13 = VB6.Format(oRecordSet.Fields.Item("U_JBTH13").Value, new string("0", Strings.Len(Drec.BIGH13)));
//            Drec.BIGI01 = VB6.Format(oRecordSet.Fields.Item("U_JBTI01").Value, new string("0", Strings.Len(Drec.BIGI01)));
//            Drec.BIGK01 = VB6.Format(oRecordSet.Fields.Item("U_JBTK01").Value, new string("0", Strings.Len(Drec.BIGK01)));
//            Drec.BIGM01 = VB6.Format(oRecordSet.Fields.Item("U_JBTM01").Value, new string("0", Strings.Len(Drec.BIGM01)));
//            Drec.BIGM02 = VB6.Format(oRecordSet.Fields.Item("U_JBTM02").Value, new string("0", Strings.Len(Drec.BIGM02)));
//            Drec.BIGM03 = VB6.Format(oRecordSet.Fields.Item("U_JBTM03").Value, new string("0", Strings.Len(Drec.BIGM03)));
//            Drec.BIGO01 = VB6.Format(oRecordSet.Fields.Item("U_JBTO01").Value, new string("0", Strings.Len(Drec.BIGO01)));
//            Drec.BIGQ01 = VB6.Format(oRecordSet.Fields.Item("U_JBTQ01").Value, new string("0", Strings.Len(Drec.BIGQ01)));
//            Drec.BIGR10 = VB6.Format(oRecordSet.Fields.Item("U_JBTY01").Value, new string("0", Strings.Len(Drec.BIGR10)));
//            Drec.BIGS01 = VB6.Format(oRecordSet.Fields.Item("U_JBTS01").Value, new string("0", Strings.Len(Drec.BIGS01)));
//            Drec.BIGT01 = VB6.Format(oRecordSet.Fields.Item("U_JBTT01").Value, new string("0", Strings.Len(Drec.BIGT01)));
//            // Drec.BIGX01 = Format$(oRecordSet.Fields("U_JBTX01").Value, String$(Len(Drec.BIGX01), "0"))
//            Drec.BIGY02 = VB6.Format(oRecordSet.Fields.Item("U_JBTY02").Value, new string("0", Strings.Len(Drec.BIGY02)));
//            Drec.BIGY03 = VB6.Format(oRecordSet.Fields.Item("U_JBTY03").Value, new string("0", Strings.Len(Drec.BIGY03)));
//            Drec.BIGY21 = VB6.Format(oRecordSet.Fields.Item("U_JBTY21").Value, new string("0", Strings.Len(Drec.BIGY21)));
//            Drec.BIGZ01 = VB6.Format(oRecordSet.Fields.Item("U_JBTZ01").Value, new string("0", Strings.Len(Drec.BIGZ01)));
//            Drec.BIGY22 = VB6.Format(oRecordSet.Fields.Item("U_JBTY20").Value, new string("0", Strings.Len(Drec.BIGY22)));
//            Drec.BIGTOT = VB6.Format(oRecordSet.Fields.Item("U_JBTTOT").Value - oRecordSet.Fields.Item("U_JBTT01").Value - oRecordSet.Fields.Item("U_JBTZ01").Value, new string("0", Strings.Len(Drec.BIGTOT)));
//            Drec.BIGTO1 = VB6.Format(oRecordSet.Fields.Item("U_JBTT01").Value + oRecordSet.Fields.Item("U_JBTZ01").Value, new string("0", Strings.Len(Drec.BIGTO1)));

//            Drec.NANGAB = VB6.Format(oRecordSet.Fields.Item("U_JONGAB").Value, new string("0", Strings.Len(Drec.NANGAB)));
//            Drec.NANJUM = VB6.Format(oRecordSet.Fields.Item("U_JONJUM").Value, new string("0", Strings.Len(Drec.NANJUM)));
//            Drec.NANNON = VB6.Format(oRecordSet.Fields.Item("U_JONNON").Value, new string("0", Strings.Len(Drec.NANNON)));
//            // Drec.NANTOT = Format$(oRecordSet.Fields("U_JONGAB").Value + oRecordSet.Fields("U_JONJUM").Value + oRecordSet.Fields("U_JONNON").Value, String$(Len(Drec.NANTOT), "0"))

//            Drec.JONCNT = VB6.Format(JONCNT, new string("0", Strings.Len(Drec.JONCNT)));
//            Drec.FILLER = Strings.Space(Strings.Len(Drec.FILLER));
//            PrintLine(1, MDC_SetMod.sStr(Drec.RECGBN) + MDC_SetMod.sStr(Drec.DTAGBN) + MDC_SetMod.sStr(Drec.TAXCOD) + MDC_SetMod.sStr(Drec.SQNNBR) + MDC_SetMod.sStr(Drec.BUSNBR) + MDC_SetMod.sStr(Drec.FILD01) + MDC_SetMod.sStr(Drec.PERNBR) + MDC_SetMod.sStr(Drec.TAXJOH) + MDC_SetMod.sStr(Drec.JONNAM) + MDC_SetMod.sStr(Drec.JONNBR) + MDC_SetMod.sStr(Drec.STRINT) + MDC_SetMod.sStr(Drec.ENDINT) + MDC_SetMod.sStr(Drec.STRGAM) + MDC_SetMod.sStr(Drec.ENDGAM) + MDC_SetMod.sStr(Drec.PAYAMT) + MDC_SetMod.sStr(Drec.BNSAMT) + MDC_SetMod.sStr(Drec.INJBNS) + MDC_SetMod.sStr(Drec.JUSBNS) + MDC_SetMod.sStr(Drec.URIBNS) + MDC_SetMod.sStr(Drec.FILD02) + MDC_SetMod.sStr(Drec.TOTAMT) + MDC_SetMod.sStr(Drec.BIGG01) + MDC_SetMod.sStr(Drec.BIGH01) + MDC_SetMod.sStr(Drec.BIGH05) + MDC_SetMod.sStr(Drec.BIGH06) + MDC_SetMod.sStr(Drec.BIGH07) + MDC_SetMod.sStr(Drec.BIGH08) + MDC_SetMod.sStr(Drec.BIGH09) + MDC_SetMod.sStr(Drec.BIGH10) + MDC_SetMod.sStr(Drec.BIGH11) + MDC_SetMod.sStr(Drec.BIGH12) + MDC_SetMod.sStr(Drec.BIGH13) + MDC_SetMod.sStr(Drec.BIGI01) + MDC_SetMod.sStr(Drec.BIGK01) + MDC_SetMod.sStr(Drec.BIGM01) + MDC_SetMod.sStr(Drec.BIGM02) + MDC_SetMod.sStr(Drec.BIGM03) + MDC_SetMod.sStr(Drec.BIGO01) + MDC_SetMod.sStr(Drec.BIGQ01) + MDC_SetMod.sStr(Drec.BIGR10) + MDC_SetMod.sStr(Drec.BIGS01) + MDC_SetMod.sStr(Drec.BIGT01) + MDC_SetMod.sStr(Drec.BIGY02) + MDC_SetMod.sStr(Drec.BIGY03) + MDC_SetMod.sStr(Drec.BIGY21) + MDC_SetMod.sStr(Drec.BIGZ01) + MDC_SetMod.sStr(Drec.BIGY22) + MDC_SetMod.sStr(Drec.BIGTOT) + MDC_SetMod.sStr(Drec.BIGTO1) + MDC_SetMod.sStr(Drec.NANGAB) + MDC_SetMod.sStr(Drec.NANJUM) + MDC_SetMod.sStr(Drec.NANNON) + MDC_SetMod.sStr(Drec.JONCNT) + MDC_SetMod.sStr(Drec.FILLER));

//            // / 필수입력 체크
//            if (Strings.Trim(Drec.JONNAM) == "")
//            {
//                Matrix_AddRow("D레코드:종전근무처 법인명(상호)가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckD = System.Convert.ToString(true);
//            }
//            if (Strings.Trim(Drec.JONNBR) == "")
//            {
//                Matrix_AddRow("D레코드:종전근무처 사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                CheckD = System.Convert.ToString(true);
//            }

//            oRecordSet.MoveNext();
//        }

//        Matrix_AddRow("D레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 종전근무지 생성 완료.", ref true);

//        if (System.Convert.ToBoolean(CheckD) == false)
//            File_Create_DRecord = true;
//        else
//            File_Create_DRecord = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("종전근무지레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("D레코드오류: " + Information.Err.Description, ref false);
//        File_Create_DRecord = false;
//    }

//    private bool File_Create_ERecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 169547
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string CheckE;
//        short BUYCNT;
//        short FAMCNT;
//        short i;

//        double TOTBOH;
//        double TOTMED;
//        double TOTEDC;
//        double TOTCAD;
//        double TOTCA1;
//        double TOTCSH;
//        double TOTGBU;
//        double TOTAMT;
//        CheckE = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        E_BUYCNT = System.Convert.ToString(0);
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / 부양가족정보
//        sQry = " Exec RPY504_2  '" + oJsnYear + "', '01', '12', '3', N'" + C_CLTCOD + "', N'%',  N'%','" + C_MSTCOD + "'";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount > 0)
//        {
//            // / E Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            Erec.RECGBN = "E";
//            Erec.DTAGBN = "20";
//            Erec.TAXCOD = Arec.TAXCOD;
//            Erec.SQNNBR = VB6.Format(NEWCNT, new string("0", Strings.Len(Erec.SQNNBR)));
//            Erec.BUSNBR = Brec.BUSNBR;
//            Erec.PERNBR = Crec.PERNBR;
//            Erec.FILLER = Strings.Space(Strings.Len(Erec.FILLER));

//            TOTBOH = 0; TOTMED = 0; TOTEDC = 0; TOTCAD = 0; TOTCSH = 0; TOTGBU = 0;

//            BUYCNT = 0;
//            FAMCNT = 1;
//            while (!oRecordSet.EOF)
//            {
//                TOTAMT = 0;
//                TOTAMT = Val(oRecordSet.Fields.Item("U_BOHAMT1").Value) + Val(oRecordSet.Fields.Item("U_BOHAMT2").Value) + Val(oRecordSet.Fields.Item("U_EDCAMT1").Value) + Val(oRecordSet.Fields.Item("U_EDCAMT2").Value) + Val(oRecordSet.Fields.Item("U_MEDAMT1").Value) + Val(oRecordSet.Fields.Item("U_MEDAMT2").Value) + Val(oRecordSet.Fields.Item("U_CADAMT1").Value) + Val(oRecordSet.Fields.Item("U_CADAMT2").Value) + Val(oRecordSet.Fields.Item("U_CSHCAD1").Value) + Val(oRecordSet.Fields.Item("U_CSHCAD2").Value) + Val(oRecordSet.Fields.Item("U_CSHAMT1").Value) + Val(oRecordSet.Fields.Item("U_GBUAMT1").Value) + Val(oRecordSet.Fields.Item("U_GBUAMT2").Value);
//                if ((Val(oRecordSet.Fields.Item("U_CHKBAS").Value) == 1 | Val(oRecordSet.Fields.Item("U_CHKCHL").Value) == 1) | (TOTAMT > 0))
//                {
//                    BUYCNT = BUYCNT + 1; // / 해당사원의 부양가족일련번호
//                    if (Val(oRecordSet.Fields.Item("U_CHKBAS").Value) == 1)
//                        E_BUYCNT = System.Convert.ToString(System.Convert.ToDouble(E_BUYCNT) + 1);
//                    // /초기화
//                    if (BUYCNT == 1)
//                    {
//                        for (i = 1; i <= 5; i++)
//                        {
//                            Erec.CHKCOD[i] = Strings.Space(Strings.Len(Erec.CHKCOD[i]));
//                            Erec.CHKINT[i] = Strings.Space(Strings.Len(Erec.CHKINT[i]));
//                            Erec.CHKNAM[i] = Strings.Space(Strings.Len(Erec.CHKNAM[i]));
//                            Erec.CHKPER[i] = Strings.Space(Strings.Len(Erec.CHKPER[i]));
//                            Erec.CHKBAS[i] = Strings.Space(Strings.Len(Erec.CHKBAS[i]));
//                            Erec.CHKJAN[i] = Strings.Space(Strings.Len(Erec.CHKJAN[i]));
//                            Erec.CHKBY6[i] = Strings.Space(Strings.Len(Erec.CHKBY6[i]));
//                            Erec.CHKBUY[i] = Strings.Space(Strings.Len(Erec.CHKBUY[i]));
//                            Erec.CHKJEL[i] = Strings.Space(Strings.Len(Erec.CHKJEL[i]));
//                            // Erec.CHKDAG(i) = Space$(Len(Erec.CHKDAG(i)))
//                            Erec.CHKCHS[i] = Strings.Space(Strings.Len(Erec.CHKCHS[i]));

//                            Erec.BOHAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.BOHAM1[i])));
//                            Erec.MEDAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.MEDAM1[i])));
//                            Erec.EDCAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.EDCAM1[i])));
//                            Erec.CADAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.CADAM1[i])));
//                            Erec.CSHCA1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.CSHCA1[i])));
//                            Erec.CSHAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.CSHAM1[i])));
//                            Erec.GBUAM1[i] = VB6.Format(0, new string("0", Strings.Len(Erec.GBUAM1[i])));
//                            Erec.BOHAM2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.BOHAM2[i])));
//                            Erec.MEDAM2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.MEDAM2[i])));
//                            Erec.EDCAM2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.EDCAM2[i])));
//                            Erec.CADAM2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.CADAM2[i])));
//                            Erec.CSHCA2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.CSHCA2[i])));
//                            Erec.GBUAM2[i] = VB6.Format(0, new string("0", Strings.Len(Erec.GBUAM2[i])));
//                        }
//                    }
//                    Erec.CHKCOD[BUYCNT] = Trim(oRecordSet.Fields.Item("U_CHKCOD").Value);
//                    // UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    Erec.CHKINT[BUYCNT] = oRecordSet.Fields.Item("U_CHKINT").Value; // / 내외국인
//                    Erec.CHKNAM[BUYCNT] = Trim(oRecordSet.Fields.Item("U_FAMNAM").Value);
//                    Erec.CHKPER[BUYCNT] = Replace(oRecordSet.Fields.Item("U_FAMPER").Value, "-", "");
//                    Erec.CHKBAS[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKBAS").Value == "0", "", oRecordSet.Fields.Item("U_CHKBAS").Value);
//                    Erec.CHKJAN[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKJAN").Value == "0", "", oRecordSet.Fields.Item("U_CHKJAN").Value);
//                    Erec.CHKBY6[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKCHL").Value == "0", "", oRecordSet.Fields.Item("U_CHKCHL").Value);
//                    Erec.CHKBUY[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKBUY").Value == "0", "", oRecordSet.Fields.Item("U_CHKBUY").Value);
//                    Erec.CHKJEL[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKJEL").Value == "0", "", oRecordSet.Fields.Item("U_CHKJEL").Value);
//                    // Erec.CHKDAG(BUYCNT) = IIf(oRecordSet.Fields("U_CHKDAJ").Value = "0", "", oRecordSet.Fields("U_CHKDAJ").Value)
//                    Erec.CHKCHS[BUYCNT] = IIf(oRecordSet.Fields.Item("U_CHKCHS").Value == "0", "", oRecordSet.Fields.Item("U_CHKCHS").Value);

//                    Erec.BOHAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_BOHAMT1").Value, new string("0", Strings.Len(Erec.BOHAM1[BUYCNT])));
//                    Erec.MEDAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_MEDAMT1").Value, new string("0", Strings.Len(Erec.MEDAM1[BUYCNT])));
//                    Erec.EDCAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_EDCAMT1").Value, new string("0", Strings.Len(Erec.EDCAM1[BUYCNT])));
//                    Erec.CADAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_CADAMT1").Value, new string("0", Strings.Len(Erec.CADAM1[BUYCNT])));
//                    Erec.CSHCA1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_CSHCAD1").Value, new string("0", Strings.Len(Erec.CSHCA1[BUYCNT])));
//                    Erec.CSHAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_CSHAMT1").Value, new string("0", Strings.Len(Erec.CSHAM1[BUYCNT])));
//                    Erec.GBUAM1[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_GBUAMT1").Value, new string("0", Strings.Len(Erec.GBUAM1[BUYCNT])));
//                    Erec.BOHAM2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_BOHAMT2").Value, new string("0", Strings.Len(Erec.BOHAM2[BUYCNT])));
//                    Erec.MEDAM2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_MEDAMT2").Value, new string("0", Strings.Len(Erec.MEDAM2[BUYCNT])));
//                    Erec.EDCAM2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_EDCAMT2").Value, new string("0", Strings.Len(Erec.EDCAM2[BUYCNT])));
//                    Erec.CADAM2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_CADAMT2").Value, new string("0", Strings.Len(Erec.CADAM2[BUYCNT])));
//                    Erec.CSHCA2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_CSHCAD2").Value, new string("0", Strings.Len(Erec.CSHCA2[BUYCNT])));
//                    Erec.GBUAM2[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("U_GBUAMT2").Value, new string("0", Strings.Len(Erec.GBUAM2[BUYCNT])));
//                    // / 부양가족명세서상 총계
//                    TOTBOH = TOTBOH + Val(oRecordSet.Fields.Item("U_BOHAMT1").Value) + Val(oRecordSet.Fields.Item("U_BOHAMT2").Value);
//                    TOTMED = TOTMED + Val(oRecordSet.Fields.Item("U_MEDAMT1").Value) + Val(oRecordSet.Fields.Item("U_MEDAMT2").Value);
//                    TOTEDC = TOTEDC + Val(oRecordSet.Fields.Item("U_EDCAMT1").Value) + Val(oRecordSet.Fields.Item("U_EDCAMT2").Value);
//                    TOTCAD = TOTCAD + Val(oRecordSet.Fields.Item("U_CADAMT1").Value) + Val(oRecordSet.Fields.Item("U_CADAMT2").Value);
//                    TOTCA1 = TOTCA1 + Val(oRecordSet.Fields.Item("U_CSHCAD1").Value) + Val(oRecordSet.Fields.Item("U_CSHCAD2").Value);
//                    TOTCSH = TOTCSH + Val(oRecordSet.Fields.Item("U_CSHAMT1").Value);
//                    TOTGBU = TOTGBU + Val(oRecordSet.Fields.Item("U_GBUAMT1").Value) + Val(oRecordSet.Fields.Item("U_GBUAMT2").Value);

//                    // E_BUYCNT = E_BUYCNT + Val(Erec.CHKBAS(BUYCNT))
//                    // / 필수입력 체크
//                    if (Strings.Trim(Erec.CHKPER[BUYCNT]) == "")
//                    {
//                        Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + Strings.Trim(Erec.CHKNAM[BUYCNT]) + "주민등록번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                        CheckE = System.Convert.ToString(true);
//                    }
//                    if (Strings.Trim(Erec.CHKNAM[BUYCNT]) == "")
//                    {
//                        Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + BUYCNT + ")번째 가족의 성명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                        CheckE = System.Convert.ToString(true);
//                    }
//                    if (Strings.Trim(Erec.CHKINT[BUYCNT]) == "")
//                    {
//                        Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + Strings.Trim(Erec.CHKNAM[BUYCNT]) + "의 내외국인구분이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                        CheckE = System.Convert.ToString(true);
//                    }
//                    if (Strings.Trim(Erec.CHKCOD[BUYCNT]) == "")
//                    {
//                        Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + Strings.Trim(Erec.CHKNAM[BUYCNT]) + "의 관계코드가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                        CheckE = System.Convert.ToString(true);
//                    }
//                    // / 마지막줄이면 다음줄로
//                    if (BUYCNT == 5)
//                    {
//                        Erec.FAMNBR = VB6.Format(FAMCNT, new string("0", Strings.Len(Erec.FAMNBR)));
//                        // / E레코드삽입
//                        PrintLine(1, MDC_SetMod.sStr(Erec.RECGBN) + MDC_SetMod.sStr(Erec.DTAGBN) + MDC_SetMod.sStr(Erec.TAXCOD) + MDC_SetMod.sStr(Erec.SQNNBR) + MDC_SetMod.sStr(Erec.BUSNBR) + MDC_SetMod.sStr(Erec.PERNBR) + MDC_SetMod.sStr(Erec.CHKCOD[1]) + MDC_SetMod.sStr(Erec.CHKINT[1]) + MDC_SetMod.sStr(Erec.CHKNAM[1]) + MDC_SetMod.sStr(Erec.CHKPER[1]) + MDC_SetMod.sStr(Erec.CHKBAS[1]) + MDC_SetMod.sStr(Erec.CHKJAN[1]) + MDC_SetMod.sStr(Erec.CHKBY6[1]) + MDC_SetMod.sStr(Erec.CHKBUY[1]) + MDC_SetMod.sStr(Erec.CHKJEL[1]) + MDC_SetMod.sStr(Erec.CHKCHS[1]) + MDC_SetMod.sStr(Erec.BOHAM1[1]) + MDC_SetMod.sStr(Erec.MEDAM1[1]) + MDC_SetMod.sStr(Erec.EDCAM1[1]) + MDC_SetMod.sStr(Erec.CADAM1[1]) + MDC_SetMod.sStr(Erec.CSHCA1[1]) + MDC_SetMod.sStr(Erec.CSHAM1[1]) + MDC_SetMod.sStr(Erec.GBUAM1[1]) + MDC_SetMod.sStr(Erec.BOHAM2[1]) + MDC_SetMod.sStr(Erec.MEDAM2[1]) + MDC_SetMod.sStr(Erec.EDCAM2[1]) + MDC_SetMod.sStr(Erec.CADAM2[1]) + MDC_SetMod.sStr(Erec.CSHCA2[1]) + MDC_SetMod.sStr(Erec.GBUAM2[1]) + MDC_SetMod.sStr(Erec.CHKCOD[2]) + MDC_SetMod.sStr(Erec.CHKINT[2]) + MDC_SetMod.sStr(Erec.CHKNAM[2]) + MDC_SetMod.sStr(Erec.CHKPER[2]) + MDC_SetMod.sStr(Erec.CHKBAS[2]) + MDC_SetMod.sStr(Erec.CHKJAN[2]) + MDC_SetMod.sStr(Erec.CHKBY6[2]) + MDC_SetMod.sStr(Erec.CHKBUY[2]) + MDC_SetMod.sStr(Erec.CHKJEL[2]) + MDC_SetMod.sStr(Erec.CHKCHS[2]) + MDC_SetMod.sStr(Erec.BOHAM1[2]) + MDC_SetMod.sStr(Erec.MEDAM1[2]) + MDC_SetMod.sStr(Erec.EDCAM1[2]) + MDC_SetMod.sStr(Erec.CADAM1[2]) + MDC_SetMod.sStr(Erec.CSHCA1[2]) + MDC_SetMod.sStr(Erec.CSHAM1[2]) + MDC_SetMod.sStr(Erec.GBUAM1[2]) + MDC_SetMod.sStr(Erec.BOHAM2[2]) + MDC_SetMod.sStr(Erec.MEDAM2[2]) + MDC_SetMod.sStr(Erec.EDCAM2[2]) + MDC_SetMod.sStr(Erec.CADAM2[2]) + MDC_SetMod.sStr(Erec.CSHCA2[2]) + MDC_SetMod.sStr(Erec.GBUAM2[2]) + MDC_SetMod.sStr(Erec.CHKCOD[3]) + MDC_SetMod.sStr(Erec.CHKINT[3]) + MDC_SetMod.sStr(Erec.CHKNAM[3]) + MDC_SetMod.sStr(Erec.CHKPER[3]) + MDC_SetMod.sStr(Erec.CHKBAS[3]) + MDC_SetMod.sStr(Erec.CHKJAN[3]) + MDC_SetMod.sStr(Erec.CHKBY6[3]) + MDC_SetMod.sStr(Erec.CHKBUY[3]) + MDC_SetMod.sStr(Erec.CHKJEL[3]) + MDC_SetMod.sStr(Erec.CHKCHS[3]) + MDC_SetMod.sStr(Erec.BOHAM1[3]) + MDC_SetMod.sStr(Erec.MEDAM1[3]) + MDC_SetMod.sStr(Erec.EDCAM1[3]) + MDC_SetMod.sStr(Erec.CADAM1[3]) + MDC_SetMod.sStr(Erec.CSHCA1[3]) + MDC_SetMod.sStr(Erec.CSHAM1[3]) + MDC_SetMod.sStr(Erec.GBUAM1[3]) + MDC_SetMod.sStr(Erec.BOHAM2[3]) + MDC_SetMod.sStr(Erec.MEDAM2[3]) + MDC_SetMod.sStr(Erec.EDCAM2[3]) + MDC_SetMod.sStr(Erec.CADAM2[3]) + MDC_SetMod.sStr(Erec.CSHCA2[3]) + MDC_SetMod.sStr(Erec.GBUAM2[3]) + MDC_SetMod.sStr(Erec.CHKCOD[4]) + MDC_SetMod.sStr(Erec.CHKINT[4]) + MDC_SetMod.sStr(Erec.CHKNAM[4]) + MDC_SetMod.sStr(Erec.CHKPER[4]) + MDC_SetMod.sStr(Erec.CHKBAS[4]) + MDC_SetMod.sStr(Erec.CHKJAN[4]) + MDC_SetMod.sStr(Erec.CHKBY6[4]) + MDC_SetMod.sStr(Erec.CHKBUY[4]) + MDC_SetMod.sStr(Erec.CHKJEL[4]) + MDC_SetMod.sStr(Erec.CHKCHS[4]) + MDC_SetMod.sStr(Erec.BOHAM1[4]) + MDC_SetMod.sStr(Erec.MEDAM1[4]) + MDC_SetMod.sStr(Erec.EDCAM1[4]) + MDC_SetMod.sStr(Erec.CADAM1[4]) + MDC_SetMod.sStr(Erec.CSHCA1[4]) + MDC_SetMod.sStr(Erec.CSHAM1[4]) + MDC_SetMod.sStr(Erec.GBUAM1[4]) + MDC_SetMod.sStr(Erec.BOHAM2[4]) + MDC_SetMod.sStr(Erec.MEDAM2[4]) + MDC_SetMod.sStr(Erec.EDCAM2[4]) + MDC_SetMod.sStr(Erec.CADAM2[4]) + MDC_SetMod.sStr(Erec.CSHCA2[4]) + MDC_SetMod.sStr(Erec.GBUAM2[4]) + MDC_SetMod.sStr(Erec.CHKCOD[5]) + MDC_SetMod.sStr(Erec.CHKINT[5]) + MDC_SetMod.sStr(Erec.CHKNAM[5]) + MDC_SetMod.sStr(Erec.CHKPER[5]) + MDC_SetMod.sStr(Erec.CHKBAS[5]) + MDC_SetMod.sStr(Erec.CHKJAN[5]) + MDC_SetMod.sStr(Erec.CHKBY6[5]) + MDC_SetMod.sStr(Erec.CHKBUY[5]) + MDC_SetMod.sStr(Erec.CHKJEL[5]) + MDC_SetMod.sStr(Erec.CHKCHS[5]) + MDC_SetMod.sStr(Erec.BOHAM1[5]) + MDC_SetMod.sStr(Erec.MEDAM1[5]) + MDC_SetMod.sStr(Erec.EDCAM1[5]) + MDC_SetMod.sStr(Erec.CADAM1[5]) + MDC_SetMod.sStr(Erec.CSHCA1[5]) + MDC_SetMod.sStr(Erec.CSHAM1[5]) + MDC_SetMod.sStr(Erec.GBUAM1[5]) + MDC_SetMod.sStr(Erec.BOHAM2[5]) + MDC_SetMod.sStr(Erec.MEDAM2[5]) + MDC_SetMod.sStr(Erec.EDCAM2[5]) + MDC_SetMod.sStr(Erec.CADAM2[5]) + MDC_SetMod.sStr(Erec.CSHCA2[5]) + MDC_SetMod.sStr(Erec.GBUAM2[5]) + MDC_SetMod.sStr(Erec.FAMNBR) + MDC_SetMod.sStr(Erec.FILLER));
//                        // / 다음줄넘김
//                        BUYCNT = 0;
//                        FAMCNT = FAMCNT + 1;
//                    }
//                }
//                oRecordSet.MoveNext();
//            }
//            if (BUYCNT > 0)
//            {
//                Erec.FAMNBR = VB6.Format(FAMCNT, new string("0", Strings.Len(Erec.FAMNBR)));
//                // / E레코드삽입
//                PrintLine(1, MDC_SetMod.sStr(Erec.RECGBN) + MDC_SetMod.sStr(Erec.DTAGBN) + MDC_SetMod.sStr(Erec.TAXCOD) + MDC_SetMod.sStr(Erec.SQNNBR) + MDC_SetMod.sStr(Erec.BUSNBR) + MDC_SetMod.sStr(Erec.PERNBR) + MDC_SetMod.sStr(Erec.CHKCOD[1]) + MDC_SetMod.sStr(Erec.CHKINT[1]) + MDC_SetMod.sStr(Erec.CHKNAM[1]) + MDC_SetMod.sStr(Erec.CHKPER[1]) + MDC_SetMod.sStr(Erec.CHKBAS[1]) + MDC_SetMod.sStr(Erec.CHKJAN[1]) + MDC_SetMod.sStr(Erec.CHKBY6[1]) + MDC_SetMod.sStr(Erec.CHKBUY[1]) + MDC_SetMod.sStr(Erec.CHKJEL[1]) + MDC_SetMod.sStr(Erec.CHKCHS[1]) + MDC_SetMod.sStr(Erec.BOHAM1[1]) + MDC_SetMod.sStr(Erec.MEDAM1[1]) + MDC_SetMod.sStr(Erec.EDCAM1[1]) + MDC_SetMod.sStr(Erec.CADAM1[1]) + MDC_SetMod.sStr(Erec.CSHCA1[1]) + MDC_SetMod.sStr(Erec.CSHAM1[1]) + MDC_SetMod.sStr(Erec.GBUAM1[1]) + MDC_SetMod.sStr(Erec.BOHAM2[1]) + MDC_SetMod.sStr(Erec.MEDAM2[1]) + MDC_SetMod.sStr(Erec.EDCAM2[1]) + MDC_SetMod.sStr(Erec.CADAM2[1]) + MDC_SetMod.sStr(Erec.CSHCA2[1]) + MDC_SetMod.sStr(Erec.GBUAM2[1]) + MDC_SetMod.sStr(Erec.CHKCOD[2]) + MDC_SetMod.sStr(Erec.CHKINT[2]) + MDC_SetMod.sStr(Erec.CHKNAM[2]) + MDC_SetMod.sStr(Erec.CHKPER[2]) + MDC_SetMod.sStr(Erec.CHKBAS[2]) + MDC_SetMod.sStr(Erec.CHKJAN[2]) + MDC_SetMod.sStr(Erec.CHKBY6[2]) + MDC_SetMod.sStr(Erec.CHKBUY[2]) + MDC_SetMod.sStr(Erec.CHKJEL[2]) + MDC_SetMod.sStr(Erec.CHKCHS[2]) + MDC_SetMod.sStr(Erec.BOHAM1[2]) + MDC_SetMod.sStr(Erec.MEDAM1[2]) + MDC_SetMod.sStr(Erec.EDCAM1[2]) + MDC_SetMod.sStr(Erec.CADAM1[2]) + MDC_SetMod.sStr(Erec.CSHCA1[2]) + MDC_SetMod.sStr(Erec.CSHAM1[2]) + MDC_SetMod.sStr(Erec.GBUAM1[2]) + MDC_SetMod.sStr(Erec.BOHAM2[2]) + MDC_SetMod.sStr(Erec.MEDAM2[2]) + MDC_SetMod.sStr(Erec.EDCAM2[2]) + MDC_SetMod.sStr(Erec.CADAM2[2]) + MDC_SetMod.sStr(Erec.CSHCA2[2]) + MDC_SetMod.sStr(Erec.GBUAM2[2]) + MDC_SetMod.sStr(Erec.CHKCOD[3]) + MDC_SetMod.sStr(Erec.CHKINT[3]) + MDC_SetMod.sStr(Erec.CHKNAM[3]) + MDC_SetMod.sStr(Erec.CHKPER[3]) + MDC_SetMod.sStr(Erec.CHKBAS[3]) + MDC_SetMod.sStr(Erec.CHKJAN[3]) + MDC_SetMod.sStr(Erec.CHKBY6[3]) + MDC_SetMod.sStr(Erec.CHKBUY[3]) + MDC_SetMod.sStr(Erec.CHKJEL[3]) + MDC_SetMod.sStr(Erec.CHKCHS[3]) + MDC_SetMod.sStr(Erec.BOHAM1[3]) + MDC_SetMod.sStr(Erec.MEDAM1[3]) + MDC_SetMod.sStr(Erec.EDCAM1[3]) + MDC_SetMod.sStr(Erec.CADAM1[3]) + MDC_SetMod.sStr(Erec.CSHCA1[3]) + MDC_SetMod.sStr(Erec.CSHAM1[3]) + MDC_SetMod.sStr(Erec.GBUAM1[3]) + MDC_SetMod.sStr(Erec.BOHAM2[3]) + MDC_SetMod.sStr(Erec.MEDAM2[3]) + MDC_SetMod.sStr(Erec.EDCAM2[3]) + MDC_SetMod.sStr(Erec.CADAM2[3]) + MDC_SetMod.sStr(Erec.CSHCA2[3]) + MDC_SetMod.sStr(Erec.GBUAM2[3]) + MDC_SetMod.sStr(Erec.CHKCOD[4]) + MDC_SetMod.sStr(Erec.CHKINT[4]) + MDC_SetMod.sStr(Erec.CHKNAM[4]) + MDC_SetMod.sStr(Erec.CHKPER[4]) + MDC_SetMod.sStr(Erec.CHKBAS[4]) + MDC_SetMod.sStr(Erec.CHKJAN[4]) + MDC_SetMod.sStr(Erec.CHKBY6[4]) + MDC_SetMod.sStr(Erec.CHKBUY[4]) + MDC_SetMod.sStr(Erec.CHKJEL[4]) + MDC_SetMod.sStr(Erec.CHKCHS[4]) + MDC_SetMod.sStr(Erec.BOHAM1[4]) + MDC_SetMod.sStr(Erec.MEDAM1[4]) + MDC_SetMod.sStr(Erec.EDCAM1[4]) + MDC_SetMod.sStr(Erec.CADAM1[4]) + MDC_SetMod.sStr(Erec.CSHCA1[4]) + MDC_SetMod.sStr(Erec.CSHAM1[4]) + MDC_SetMod.sStr(Erec.GBUAM1[4]) + MDC_SetMod.sStr(Erec.BOHAM2[4]) + MDC_SetMod.sStr(Erec.MEDAM2[4]) + MDC_SetMod.sStr(Erec.EDCAM2[4]) + MDC_SetMod.sStr(Erec.CADAM2[4]) + MDC_SetMod.sStr(Erec.CSHCA2[4]) + MDC_SetMod.sStr(Erec.GBUAM2[4]) + MDC_SetMod.sStr(Erec.CHKCOD[5]) + MDC_SetMod.sStr(Erec.CHKINT[5]) + MDC_SetMod.sStr(Erec.CHKNAM[5]) + MDC_SetMod.sStr(Erec.CHKPER[5]) + MDC_SetMod.sStr(Erec.CHKBAS[5]) + MDC_SetMod.sStr(Erec.CHKJAN[5]) + MDC_SetMod.sStr(Erec.CHKBY6[5]) + MDC_SetMod.sStr(Erec.CHKBUY[5]) + MDC_SetMod.sStr(Erec.CHKJEL[5]) + MDC_SetMod.sStr(Erec.CHKCHS[5]) + MDC_SetMod.sStr(Erec.BOHAM1[5]) + MDC_SetMod.sStr(Erec.MEDAM1[5]) + MDC_SetMod.sStr(Erec.EDCAM1[5]) + MDC_SetMod.sStr(Erec.CADAM1[5]) + MDC_SetMod.sStr(Erec.CSHCA1[5]) + MDC_SetMod.sStr(Erec.CSHAM1[5]) + MDC_SetMod.sStr(Erec.GBUAM1[5]) + MDC_SetMod.sStr(Erec.BOHAM2[5]) + MDC_SetMod.sStr(Erec.MEDAM2[5]) + MDC_SetMod.sStr(Erec.EDCAM2[5]) + MDC_SetMod.sStr(Erec.CADAM2[5]) + MDC_SetMod.sStr(Erec.CSHCA2[5]) + MDC_SetMod.sStr(Erec.GBUAM2[5]) + MDC_SetMod.sStr(Erec.FAMNBR) + MDC_SetMod.sStr(Erec.FILLER));
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 부양가족레코드 생성 완료.", ref true);
//            }
//        }
//        // / 기본부양가족공제 기본공제가족수와 가족명세서의 기본공제가족수가 다를경우
//        if ((System.Convert.ToDouble(E_BUYCNT) - 1) != System.Convert.ToDouble(C_BUYCNT))
//        {
//            Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 기본공제받은 부양가족이 누락되었습니다.", ref true, ref true);
//            CheckE = System.Convert.ToString(false);
//        }
//        // / 기본부양가족공제의 금액과 소득공제항목등록한 총공제금액이 다를경우
//        sQry = "SELECT  SUM(ISNULL(U_BOHAMT,0)+ISNULL(U_JGABOA,0) + ISNULL(U_BOHAL1,0) + ISNULL(U_BOHAL2,0)) AS BOHAMT,";
//        sQry = sQry + " SUM(ISNULL(U_JGAMED,0)+ISNULL(U_GENMED,0)) AS MEDAMT,";
//        sQry = sQry + " SUM(ISNULL(U_BONSCH,0)+ISNULL(U_JGASCH,0)+ISNULL(U_JICSCH,0)) AS EDCAMT,";
//        sQry = sQry + " SUM(ISNULL(U_CADSAV,0)+ISNULL(U_GIRSAV,0)) AS CADAMT,";
//        sQry = sQry + " SUM(ISNULL(U_CA1SAV,0)) AS CA1SAV,";
//        sQry = sQry + " SUM(ISNULL(U_CSHSAV,0)) AS CSHAMT,";
//        sQry = sQry + " SUM(IsNull(U_LAWGBU, 0) + IsNull(U_POCGBU, 0) + IsNull(U_SP1GBU, 0) + IsNull(U_SP2GBU, 0) + IsNull(U_USJGBU, 0) + IsNull(U_JIJGBU, 0)++ IsNull(U_JNGGBU, 0)) As GBUAMT";
//        sQry = sQry + " FROM [@ZPY501H] WHERE U_JSNYER = '" + oJsnYear + "' AND U_MSTCOD = '" + C_MSTCOD + "'";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount > 0)
//        {
//            // /1. 보험료총액이 다를경우
//            if (TOTBOH < Val(oRecordSet.Fields.Item("BOHAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]이 부양가족명세의 보험료보다 보험료총계가 큽니다. 확인하십시오.", ref true, ref true);
//            // /2. 의료비총액이 다를경우
//            if (TOTMED != Val(oRecordSet.Fields.Item("MEDAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]과 부양가족명세의 의료비총계가 다릅니다. 확인하십시오.", ref true, ref true);
//            // /3. 교육비총액이 다를경우
//            if (TOTEDC < Val(oRecordSet.Fields.Item("EDCAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]이 부양가족명세의 교육비보다 교육비총계가 큽니다. 확인하십시오.", ref true, ref true);
//            // /4. 신용카드총액이 다를경우
//            if (TOTCAD != Val(oRecordSet.Fields.Item("CADAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]과 부양가족명세의 신용카드등총계가 다릅니다. 확인하십시오.", ref true, ref true);
//            // /5. 직불카드총액이 다를경우
//            if (TOTCA1 != Val(oRecordSet.Fields.Item("CA1SAV").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]과 부양가족명세의 직불카드총계가 다릅니다. 확인하십시오.", ref true, ref true);
//            // /6. 현금영수증총액이 다를경우
//            if (TOTCSH != Val(oRecordSet.Fields.Item("CSHAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]과 부양가족명세의 현금영수증총계가 다릅니다. 확인하십시오.", ref true, ref true);
//            // /7. 기부금총액이 다를경우
//            if (TOTGBU < Val(oRecordSet.Fields.Item("GBUAMT").Value))
//                Matrix_AddRow("E레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " [소득공제항목]이 부양가족명세의 기부금총계보다 큽니다. 확인하십시오.", ref true, ref true);
//        }
//        if (System.Convert.ToBoolean(CheckE) == false)
//            File_Create_ERecord = true;
//        else
//            File_Create_ERecord = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("부양가족레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("E레코드오류: " + Information.Err.Source + " " + Information.Err.Description, ref false);
//        File_Create_ERecord = false;
//    }

//    private bool File_Create_FRecord()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 190530
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		short ErrNum;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;
//        string CheckF;
//        short SAVCNT;
//        int iRow;

//        double RETSAV; double RETSA1; double GYNSAV; double YUNSAV;
//        double HUSAMT; double HU1AMT; double HU2AMT; double HU3AMT;
//        double JFDAM1; double JFDAM2; double JFDAM3;

//        CheckF = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / 부양가족정보
//        sQry = "EXEC ZPY521 'F', '" + oJsnYear + "', '" + JSNGBN + "', '" + STRMON + "', '" + ENDMON + "', " + "'" + C_CLTCOD + "', '" + MSTBRK + "', '" + DPTSTR + "', '" + DPTEND + "','" + C_MSTCOD + "'";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount > 0)
//        {
//            // / F Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//            Frec.RECGBN = "F";
//            Frec.DTAGBN = "20";
//            Frec.TAXCOD = Arec.TAXCOD;
//            Frec.SQNNBR = VB6.Format(NEWCNT, new string("0", Strings.Len(Frec.SQNNBR)));
//            Frec.BUSNBR = Brec.BUSNBR;
//            Frec.PERNBR = Crec.PERNBR;
//            Frec.FILLER = Strings.Space(Strings.Len(Frec.FILLER));

//            SAVCNT = 0;
//            while (!oRecordSet.EOF)
//            {
//                SAVCNT = SAVCNT + 1; // / 해당사원의 부양가족일련번호
//                                     // /초기화
//                if (SAVCNT == 1)
//                {
//                    for (iRow = 1; iRow <= 15; iRow++)
//                    {
//                        Frec.SAVGBN[iRow] = Strings.Space(Strings.Len(Frec.SAVGBN[iRow]));
//                        Frec.SAVCOD[iRow] = Strings.Space(Strings.Len(Frec.SAVCOD[iRow]));
//                        Frec.SAVNAM[iRow] = Strings.Space(Strings.Len(Frec.SAVNAM[iRow]));
//                        Frec.SAVNUM[iRow] = Strings.Space(Strings.Len(Frec.SAVNUM[iRow]));

//                        Frec.STYEAR[iRow] = VB6.Format(0, new string("0", Strings.Len(Frec.STYEAR[iRow])));
//                        Frec.SAVAMT[iRow] = VB6.Format(0, new string("0", Strings.Len(Frec.SAVAMT[iRow])));
//                        Frec.SARAMT[iRow] = VB6.Format(0, new string("0", Strings.Len(Frec.SARAMT[iRow])));
//                    }
//                }

//                Frec.SAVGBN[SAVCNT] = Trim(oRecordSet.Fields.Item("U_SAVGBN").Value);
//                Frec.SAVCOD[SAVCNT] = Trim(oRecordSet.Fields.Item("U_SAVCOD").Value);
//                Frec.SAVNAM[SAVCNT] = Trim(oRecordSet.Fields.Item("U_SAVNAM").Value);
//                Frec.SAVNUM[SAVCNT] = Trim(oRecordSet.Fields.Item("U_SAVNUM").Value);

//                Frec.STYEAR[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("U_STYEAR").Value, new string("0", Strings.Len(Frec.STYEAR[SAVCNT])));
//                Frec.SAVAMT[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("U_SAVAMT").Value, new string("0", Strings.Len(Frec.SAVAMT[SAVCNT])));
//                Frec.SARAMT[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("U_SARAMT").Value, new string("0", Strings.Len(Frec.SARAMT[SAVCNT])));

//                // / 부양가족명세서상 총계
//                switch (Trim(oRecordSet.Fields.Item("U_SAVGBN").Value))
//                {
//                    case "11":
//                        {
//                            RETSAV = RETSAV + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "12":
//                        {
//                            RETSA1 = RETSA1 + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "21":
//                        {
//                            GYNSAV = GYNSAV + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "22":
//                        {
//                            YUNSAV = YUNSAV + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "31":
//                        {
//                            HUSAMT = HUSAMT + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "32":
//                        {
//                            HU1AMT = HU1AMT + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "33":
//                        {
//                            HU2AMT = HU2AMT + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "34":
//                        {
//                            HU3AMT = HU3AMT + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                            break;
//                        }

//                    case "41":
//                        {
//                            switch (oRecordSet.Fields.Item("U_STYEAR").Value)
//                            {
//                                case "01":
//                                    {
//                                        JFDAM1 = JFDAM1 + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                                        break;
//                                    }

//                                case "02":
//                                    {
//                                        JFDAM2 = JFDAM2 + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                                        break;
//                                    }

//                                case "03":
//                                    {
//                                        JFDAM3 = JFDAM3 + oRecordSet.Fields.Item("U_SAVAMT").Value;
//                                        break;
//                                    }
//                            }

//                            break;
//                        }
//                }

//                // / 필수입력 체크
//                if (Strings.Trim(Frec.SAVGBN[SAVCNT]) == "41" & Strings.Trim(Frec.STYEAR[SAVCNT]) == "00")
//                {
//                    Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + Strings.Trim(Frec.SAVNAM[SAVCNT]) + " 납입년차가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//                    CheckF = System.Convert.ToString(true);
//                }

//                // / 마지막줄이면 다음줄로
//                if (SAVCNT == 15)
//                {
//                    // / E레코드삽입
//                    PrintLine(1, MDC_SetMod.sStr(Frec.RECGBN) + MDC_SetMod.sStr(Frec.DTAGBN) + MDC_SetMod.sStr(Frec.TAXCOD) + MDC_SetMod.sStr(Frec.SQNNBR) + MDC_SetMod.sStr(Frec.BUSNBR) + MDC_SetMod.sStr(Frec.PERNBR) + MDC_SetMod.sStr(Frec.SAVGBN[1]) + MDC_SetMod.sStr(Frec.SAVCOD[1]) + MDC_SetMod.sStr(Frec.SAVNAM[1]) + MDC_SetMod.sStr(Frec.SAVNUM[1]) + MDC_SetMod.sStr(Frec.STYEAR[1]) + MDC_SetMod.sStr(Frec.SAVAMT[1]) + MDC_SetMod.sStr(Frec.SARAMT[1]) + MDC_SetMod.sStr(Frec.SAVGBN[2]) + MDC_SetMod.sStr(Frec.SAVCOD[2]) + MDC_SetMod.sStr(Frec.SAVNAM[2]) + MDC_SetMod.sStr(Frec.SAVNUM[2]) + MDC_SetMod.sStr(Frec.STYEAR[2]) + MDC_SetMod.sStr(Frec.SAVAMT[2]) + MDC_SetMod.sStr(Frec.SARAMT[2]) + MDC_SetMod.sStr(Frec.SAVGBN[3]) + MDC_SetMod.sStr(Frec.SAVCOD[3]) + MDC_SetMod.sStr(Frec.SAVNAM[3]) + MDC_SetMod.sStr(Frec.SAVNUM[3]) + MDC_SetMod.sStr(Frec.STYEAR[3]) + MDC_SetMod.sStr(Frec.SAVAMT[3]) + MDC_SetMod.sStr(Frec.SARAMT[3]) + MDC_SetMod.sStr(Frec.SAVGBN[4]) + MDC_SetMod.sStr(Frec.SAVCOD[4]) + MDC_SetMod.sStr(Frec.SAVNAM[4]) + MDC_SetMod.sStr(Frec.SAVNUM[4]) + MDC_SetMod.sStr(Frec.STYEAR[4]) + MDC_SetMod.sStr(Frec.SAVAMT[4]) + MDC_SetMod.sStr(Frec.SARAMT[4]) + MDC_SetMod.sStr(Frec.SAVGBN[5]) + MDC_SetMod.sStr(Frec.SAVCOD[5]) + MDC_SetMod.sStr(Frec.SAVNAM[5]) + MDC_SetMod.sStr(Frec.SAVNUM[5]) + MDC_SetMod.sStr(Frec.STYEAR[5]) + MDC_SetMod.sStr(Frec.SAVAMT[5]) + MDC_SetMod.sStr(Frec.SARAMT[5]) + MDC_SetMod.sStr(Frec.SAVGBN[6]) + MDC_SetMod.sStr(Frec.SAVCOD[6]) + MDC_SetMod.sStr(Frec.SAVNAM[6]) + MDC_SetMod.sStr(Frec.SAVNUM[6]) + MDC_SetMod.sStr(Frec.STYEAR[6]) + MDC_SetMod.sStr(Frec.SAVAMT[6]) + MDC_SetMod.sStr(Frec.SARAMT[6]) + MDC_SetMod.sStr(Frec.SAVGBN[7]) + MDC_SetMod.sStr(Frec.SAVCOD[7]) + MDC_SetMod.sStr(Frec.SAVNAM[7]) + MDC_SetMod.sStr(Frec.SAVNUM[7]) + MDC_SetMod.sStr(Frec.STYEAR[7]) + MDC_SetMod.sStr(Frec.SAVAMT[7]) + MDC_SetMod.sStr(Frec.SARAMT[7]) + MDC_SetMod.sStr(Frec.SAVGBN[8]) + MDC_SetMod.sStr(Frec.SAVCOD[8]) + MDC_SetMod.sStr(Frec.SAVNAM[8]) + MDC_SetMod.sStr(Frec.SAVNUM[8]) + MDC_SetMod.sStr(Frec.STYEAR[8]) + MDC_SetMod.sStr(Frec.SAVAMT[8]) + MDC_SetMod.sStr(Frec.SARAMT[8]) + MDC_SetMod.sStr(Frec.SAVGBN[9]) + MDC_SetMod.sStr(Frec.SAVCOD[9]) + MDC_SetMod.sStr(Frec.SAVNAM[9]) + MDC_SetMod.sStr(Frec.SAVNUM[9]) + MDC_SetMod.sStr(Frec.STYEAR[9]) + MDC_SetMod.sStr(Frec.SAVAMT[9]) + MDC_SetMod.sStr(Frec.SARAMT[9]) + MDC_SetMod.sStr(Frec.SAVGBN[10]) + MDC_SetMod.sStr(Frec.SAVCOD[10]) + MDC_SetMod.sStr(Frec.SAVNAM[10]) + MDC_SetMod.sStr(Frec.SAVNUM[10]) + MDC_SetMod.sStr(Frec.STYEAR[10]) + MDC_SetMod.sStr(Frec.SAVAMT[10]) + MDC_SetMod.sStr(Frec.SARAMT[10]) + MDC_SetMod.sStr(Frec.SAVGBN[11]) + MDC_SetMod.sStr(Frec.SAVCOD[11]) + MDC_SetMod.sStr(Frec.SAVNAM[11]) + MDC_SetMod.sStr(Frec.SAVNUM[11]) + MDC_SetMod.sStr(Frec.STYEAR[11]) + MDC_SetMod.sStr(Frec.SAVAMT[11]) + MDC_SetMod.sStr(Frec.SARAMT[11]) + MDC_SetMod.sStr(Frec.SAVGBN[12]) + MDC_SetMod.sStr(Frec.SAVCOD[12]) + MDC_SetMod.sStr(Frec.SAVNAM[12]) + MDC_SetMod.sStr(Frec.SAVNUM[12]) + MDC_SetMod.sStr(Frec.STYEAR[12]) + MDC_SetMod.sStr(Frec.SAVAMT[12]) + MDC_SetMod.sStr(Frec.SARAMT[12]) + MDC_SetMod.sStr(Frec.SAVGBN[13]) + MDC_SetMod.sStr(Frec.SAVCOD[13]) + MDC_SetMod.sStr(Frec.SAVNAM[13]) + MDC_SetMod.sStr(Frec.SAVNUM[13]) + MDC_SetMod.sStr(Frec.STYEAR[13]) + MDC_SetMod.sStr(Frec.SAVAMT[13]) + MDC_SetMod.sStr(Frec.SARAMT[13]) + MDC_SetMod.sStr(Frec.SAVGBN[14]) + MDC_SetMod.sStr(Frec.SAVCOD[14]) + MDC_SetMod.sStr(Frec.SAVNAM[14]) + MDC_SetMod.sStr(Frec.SAVNUM[14]) + MDC_SetMod.sStr(Frec.STYEAR[14]) + MDC_SetMod.sStr(Frec.SAVAMT[14]) + MDC_SetMod.sStr(Frec.SARAMT[14]) + MDC_SetMod.sStr(Frec.SAVGBN[15]) + MDC_SetMod.sStr(Frec.SAVCOD[15]) + MDC_SetMod.sStr(Frec.SAVNAM[15]) + MDC_SetMod.sStr(Frec.SAVNUM[15]) + MDC_SetMod.sStr(Frec.STYEAR[15]) + MDC_SetMod.sStr(Frec.SAVAMT[15]) + MDC_SetMod.sStr(Frec.SARAMT[15]) + MDC_SetMod.sStr(Frec.FILLER));
//                    // / 다음줄넘김
//                    SAVCNT = 0;
//                }

//                oRecordSet.MoveNext();
//            }
//            if (SAVCNT > 0)
//            {
//                // / E레코드삽입
//                PrintLine(1, MDC_SetMod.sStr(Frec.RECGBN) + MDC_SetMod.sStr(Frec.DTAGBN) + MDC_SetMod.sStr(Frec.TAXCOD) + MDC_SetMod.sStr(Frec.SQNNBR) + MDC_SetMod.sStr(Frec.BUSNBR) + MDC_SetMod.sStr(Frec.PERNBR) + MDC_SetMod.sStr(Frec.SAVGBN[1]) + MDC_SetMod.sStr(Frec.SAVCOD[1]) + MDC_SetMod.sStr(Frec.SAVNAM[1]) + MDC_SetMod.sStr(Frec.SAVNUM[1]) + MDC_SetMod.sStr(Frec.STYEAR[1]) + MDC_SetMod.sStr(Frec.SAVAMT[1]) + MDC_SetMod.sStr(Frec.SARAMT[1]) + MDC_SetMod.sStr(Frec.SAVGBN[2]) + MDC_SetMod.sStr(Frec.SAVCOD[2]) + MDC_SetMod.sStr(Frec.SAVNAM[2]) + MDC_SetMod.sStr(Frec.SAVNUM[2]) + MDC_SetMod.sStr(Frec.STYEAR[2]) + MDC_SetMod.sStr(Frec.SAVAMT[2]) + MDC_SetMod.sStr(Frec.SARAMT[2]) + MDC_SetMod.sStr(Frec.SAVGBN[3]) + MDC_SetMod.sStr(Frec.SAVCOD[3]) + MDC_SetMod.sStr(Frec.SAVNAM[3]) + MDC_SetMod.sStr(Frec.SAVNUM[3]) + MDC_SetMod.sStr(Frec.STYEAR[3]) + MDC_SetMod.sStr(Frec.SAVAMT[3]) + MDC_SetMod.sStr(Frec.SARAMT[3]) + MDC_SetMod.sStr(Frec.SAVGBN[4]) + MDC_SetMod.sStr(Frec.SAVCOD[4]) + MDC_SetMod.sStr(Frec.SAVNAM[4]) + MDC_SetMod.sStr(Frec.SAVNUM[4]) + MDC_SetMod.sStr(Frec.STYEAR[4]) + MDC_SetMod.sStr(Frec.SAVAMT[4]) + MDC_SetMod.sStr(Frec.SARAMT[4]) + MDC_SetMod.sStr(Frec.SAVGBN[5]) + MDC_SetMod.sStr(Frec.SAVCOD[5]) + MDC_SetMod.sStr(Frec.SAVNAM[5]) + MDC_SetMod.sStr(Frec.SAVNUM[5]) + MDC_SetMod.sStr(Frec.STYEAR[5]) + MDC_SetMod.sStr(Frec.SAVAMT[5]) + MDC_SetMod.sStr(Frec.SARAMT[5]) + MDC_SetMod.sStr(Frec.SAVGBN[6]) + MDC_SetMod.sStr(Frec.SAVCOD[6]) + MDC_SetMod.sStr(Frec.SAVNAM[6]) + MDC_SetMod.sStr(Frec.SAVNUM[6]) + MDC_SetMod.sStr(Frec.STYEAR[6]) + MDC_SetMod.sStr(Frec.SAVAMT[6]) + MDC_SetMod.sStr(Frec.SARAMT[6]) + MDC_SetMod.sStr(Frec.SAVGBN[7]) + MDC_SetMod.sStr(Frec.SAVCOD[7]) + MDC_SetMod.sStr(Frec.SAVNAM[7]) + MDC_SetMod.sStr(Frec.SAVNUM[7]) + MDC_SetMod.sStr(Frec.STYEAR[7]) + MDC_SetMod.sStr(Frec.SAVAMT[7]) + MDC_SetMod.sStr(Frec.SARAMT[7]) + MDC_SetMod.sStr(Frec.SAVGBN[8]) + MDC_SetMod.sStr(Frec.SAVCOD[8]) + MDC_SetMod.sStr(Frec.SAVNAM[8]) + MDC_SetMod.sStr(Frec.SAVNUM[8]) + MDC_SetMod.sStr(Frec.STYEAR[8]) + MDC_SetMod.sStr(Frec.SAVAMT[8]) + MDC_SetMod.sStr(Frec.SARAMT[8]) + MDC_SetMod.sStr(Frec.SAVGBN[9]) + MDC_SetMod.sStr(Frec.SAVCOD[9]) + MDC_SetMod.sStr(Frec.SAVNAM[9]) + MDC_SetMod.sStr(Frec.SAVNUM[9]) + MDC_SetMod.sStr(Frec.STYEAR[9]) + MDC_SetMod.sStr(Frec.SAVAMT[9]) + MDC_SetMod.sStr(Frec.SARAMT[9]) + MDC_SetMod.sStr(Frec.SAVGBN[10]) + MDC_SetMod.sStr(Frec.SAVCOD[10]) + MDC_SetMod.sStr(Frec.SAVNAM[10]) + MDC_SetMod.sStr(Frec.SAVNUM[10]) + MDC_SetMod.sStr(Frec.STYEAR[10]) + MDC_SetMod.sStr(Frec.SAVAMT[10]) + MDC_SetMod.sStr(Frec.SARAMT[10]) + MDC_SetMod.sStr(Frec.SAVGBN[11]) + MDC_SetMod.sStr(Frec.SAVCOD[11]) + MDC_SetMod.sStr(Frec.SAVNAM[11]) + MDC_SetMod.sStr(Frec.SAVNUM[11]) + MDC_SetMod.sStr(Frec.STYEAR[11]) + MDC_SetMod.sStr(Frec.SAVAMT[11]) + MDC_SetMod.sStr(Frec.SARAMT[11]) + MDC_SetMod.sStr(Frec.SAVGBN[12]) + MDC_SetMod.sStr(Frec.SAVCOD[12]) + MDC_SetMod.sStr(Frec.SAVNAM[12]) + MDC_SetMod.sStr(Frec.SAVNUM[12]) + MDC_SetMod.sStr(Frec.STYEAR[12]) + MDC_SetMod.sStr(Frec.SAVAMT[12]) + MDC_SetMod.sStr(Frec.SARAMT[12]) + MDC_SetMod.sStr(Frec.SAVGBN[13]) + MDC_SetMod.sStr(Frec.SAVCOD[13]) + MDC_SetMod.sStr(Frec.SAVNAM[13]) + MDC_SetMod.sStr(Frec.SAVNUM[13]) + MDC_SetMod.sStr(Frec.STYEAR[13]) + MDC_SetMod.sStr(Frec.SAVAMT[13]) + MDC_SetMod.sStr(Frec.SARAMT[13]) + MDC_SetMod.sStr(Frec.SAVGBN[14]) + MDC_SetMod.sStr(Frec.SAVCOD[14]) + MDC_SetMod.sStr(Frec.SAVNAM[14]) + MDC_SetMod.sStr(Frec.SAVNUM[14]) + MDC_SetMod.sStr(Frec.STYEAR[14]) + MDC_SetMod.sStr(Frec.SAVAMT[14]) + MDC_SetMod.sStr(Frec.SARAMT[14]) + MDC_SetMod.sStr(Frec.SAVGBN[15]) + MDC_SetMod.sStr(Frec.SAVCOD[15]) + MDC_SetMod.sStr(Frec.SAVNAM[15]) + MDC_SetMod.sStr(Frec.SAVNUM[15]) + MDC_SetMod.sStr(Frec.STYEAR[15]) + MDC_SetMod.sStr(Frec.SAVAMT[15]) + MDC_SetMod.sStr(Frec.SARAMT[15]) + MDC_SetMod.sStr(Frec.FILLER));
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 연금저축명세 생성 완료.", ref true);
//            }
//        }

//        // / 기본부양가족공제의 금액과 소득공제항목등록한 총공제금액이 다를경우
//        sQry = "SELECT  SUM(ISNULL(U_RETSAV,0)) AS RETSAV,";
//        sQry = sQry + " SUM(ISNULL(U_RETSA1,0)) AS RETSA1,";
//        sQry = sQry + " SUM(ISNULL(U_GYNSAV,0)) AS GYNSAV,";
//        sQry = sQry + " SUM(ISNULL(U_YUNSAV,0)) AS YUNSAV,";
//        sQry = sQry + " SUM(ISNULL(U_HUSAMT,0)) AS HUSAMT,";
//        sQry = sQry + " SUM(ISNULL(U_HU1AMT,0)) AS HU1AMT,";
//        sQry = sQry + " SUM(ISNULL(U_HU2AMT,0)) AS HU2AMT,";
//        sQry = sQry + " SUM(ISNULL(U_HU3AMT,0)) AS HU3AMT,";
//        sQry = sQry + " SUM(ISNULL(U_JFDAM1,0)) AS JFDAM1,";
//        sQry = sQry + " SUM(ISNULL(U_JFDAM2,0)) AS JFDAM2,";
//        sQry = sQry + " SUM(ISNULL(U_JFDAM3,0)) AS JFDAM3 ";
//        sQry = sQry + " FROM [@ZPY501H] ";
//        sQry = sQry + " WHERE U_JSNYER = '" + oJsnYear + "'";
//        sQry = sQry + " AND U_MSTCOD = '" + C_MSTCOD + "'";
//        sQry = sQry + " AND U_CLTCOD = '" + C_CLTCOD + "'";
//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount > 0)
//        {
//            // /1. 퇴직연금(근로자 퇴직급여보장법)
//            if (RETSAV < Val(oRecordSet.Fields.Item("RETSAV").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "퇴직연금(근로자 퇴직급여보장법) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /2. 퇴직연금(과학기술인공제)
//            if (RETSA1 < Val(oRecordSet.Fields.Item("RETSA1").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "퇴직연금(과학기술인공제) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /3. 연금저축(개인연금저축)
//            if (GYNSAV < Val(oRecordSet.Fields.Item("GYNSAV").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "연금저축(개인연금저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /4. 연금저축(연금저축)
//            if (YUNSAV < Val(oRecordSet.Fields.Item("YUNSAV").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "연금저축(연금저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /5. 주택마련저축(청약저축)
//            if (HUSAMT < Val(oRecordSet.Fields.Item("HUSAMT").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(청약저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /6. 주택마련저축(주택청약종합저축)
//            if (HU1AMT < Val(oRecordSet.Fields.Item("HU1AMT").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(주택청약종합저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /7. 주택마련저축(장기주택마련저축)
//            if (HU2AMT < Val(oRecordSet.Fields.Item("HU2AMT").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(장기주택마련저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /8. 주택마련저축(근로자주택마련저축)
//            if (HU3AMT < Val(oRecordSet.Fields.Item("HU3AMT").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(근로자주택마련저축) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /9. 주택마련저축(장기주식형-1년차)
//            if (JFDAM1 < Val(oRecordSet.Fields.Item("JFDAM1").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(장기주식형-1년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /10. 주택마련저축(장기주식형-2년차)
//            if (JFDAM2 < Val(oRecordSet.Fields.Item("JFDAM2").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(장기주식형-2년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//            // /11. 주택마련저축(장기주식형-3년차)
//            if (JFDAM3 < Val(oRecordSet.Fields.Item("JFDAM3").Value))
//                Matrix_AddRow("F레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "주택마련저축(장기주식형-3년차) 금액이 [소득공제항목]과 [연금저축명세]에서 금액이 맞지 않습니다. 확인하십시오.", ref true, ref true);
//        }
//        if (System.Convert.ToBoolean(CheckF) == false)
//            File_Create_FRecord = true;
//        else
//            File_Create_FRecord = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        Matrix_AddRow("F레코드오류: " + Information.Err.Source + " " + Information.Err.Description, ref false);
//        File_Create_FRecord = false;
//    }

//    private void FlushToItemValue(string oUID, ref int oRow = 0)
//    {
//        ZPAY_g_EmpID MstInfo;

//        switch (oUID)
//        {
//            case "JsnYear":
//                {
//                    // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oJsnYear = oForm.Items.Item(oUID).Specific.String;
//                    if (Strings.Trim(oJsnYear) == "")
//                        ZPAY_GBL_JSNYER.Value = oJsnYear;
//                    break;
//                }

//            case "MSTCOD":
//                {
//                    // UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (oForm.Items.Item(oUID).Specific.String == "")
//                    {
//                        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oForm.Items.Item(oUID).Specific.String = "";
//                        oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//                        oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = "";
//                    }
//                    else
//                    {
//                        // UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oForm.Items.Item(oUID).Specific.String = UCase(oForm.Items.Item(oUID).Specific.String);
//                        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        // UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items.Item(oUID).Specific.String);
//                        oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MstInfo.MSTNAM;
//                        oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = MstInfo.EmpID;
//                    }
//                    oForm.Items.Item("MSTNAM").Update();
//                    oForm.Items.Item("EmpID").Update();
//                    break;
//                }
//        }
//        oForm.Items.Item(oUID).Update();
//    }

//    private bool HeaderSpaceLineDel()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo HeaderSpaceLi...' at character 209256
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//		'저장할 데이터의 유효성을 점검한다
//		'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//		On Error GoTo HeaderSpaceLineDel

// */		short ErrNum;

//        ErrNum = 0;
//        // / 필수Check
//        // UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        if (MDC_SetMod.ChkYearMonth(oForm.Items.Item("JsnYear").Specific.String + "01") == false)
//        {
//            ErrNum = 1;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("JSNTYP").Specific.Selected == null)
//        {
//            ErrNum = 2;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("BPLId").Specific.Selected == null)
//        {
//            ErrNum = 3;
//            goto HeaderSpaceLineDel;
//        }
//        else if (Trim(oForm.Items.Item("PRTDAT").Specific.String) == "")
//        {
//            ErrNum = 4;
//            goto HeaderSpaceLineDel;
//        }
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        MSTBRK = oForm.Items.Item("BPLId").Specific.Selected.Value;
//        // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        DPTSTR = oForm.Items.Item("DptStr").Specific.Selected.Value;
//        if (DPTSTR == "-1")
//            DPTSTR = "00000001"; // 20120209 "00000001" 변경
//                                 // UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        DPTEND = oForm.Items.Item("DptEnd").Specific.Selected.Value;
//        if (DPTEND == "-1")
//            DPTEND = "ZZZZZZZZ";
//        // UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
//        if (Strings.Trim(MSTCOD) == "")
//            MSTCOD = "%";

//        HeaderSpaceLineDel = true;
//        return;
//        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        HeaderSpaceLineDel:
//        ;
//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("귀속년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("기간코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("지점은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 4)
//            Sbo_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//        HeaderSpaceLineDel = false;
//    }

//    private void Matrix_AddRow(string MatrixMsg, ref bool Insert_YN = false, ref bool MatrixErr = false)
//    {
//        if (MatrixErr == true)
//            oForm.DataSources.UserDataSources.Item("Col0").Value = "??";
//        else
//            oForm.DataSources.UserDataSources.Item("Col0").Value = "";
//        oForm.DataSources.UserDataSources.Item("Col1").Value = MatrixMsg;
//        if (Insert_YN == true)
//        {
//            oMat1.AddRow();
//            MaxRow = MaxRow + 1;
//        }
//        oMat1.SetLineData(MaxRow);
//    }
//}
