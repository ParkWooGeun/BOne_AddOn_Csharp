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
//	internal class ZPY421
//	{
//////********************************************************************************
//////  File           : ZPY421.cls
//////  Module         : 인사관리 > 퇴직소득전산매체수록
//////  Desc           : 퇴직소득전산매체수록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

////'// 그리드 사용시
////Public oGrid1           As SAPbouiCOM.Grid
////Public oDS_ZPY421     As SAPbouiCOM.DataTable
////
////'// 매트릭스 사용시
//		public SAPbouiCOM.Matrix oMat1;

////Private oDS_ZPY421A As SAPbouiCOM.DBDataSource
////Private oDS_ZPY421B As SAPbouiCOM.DBDataSource

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;


//		private string oJsnYear;
//		private string DPTSTR;
//		private string DPTEND;
//		private string MSTCOD;
//		private string CLTCOD;
//			/// B레코드일련번호
//		private short BUSCNT;
//			/// B레코드총갯수
//		private short BUSTOT;

//		private string oFilePath;

//			//파  일  명
//		private Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString FILNAM = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(30);
//		private int MaxRow;
//		private short NEWCNT;
//		private string C_MSTCOD;

///// 퇴직소득 지급명세서-1
//		private struct A_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] RECGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료  구분
//			public char[] DTAGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세  무  서
//			public char[] TAXCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//제출  일자
//			public char[] PRTDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제  출  자 (1;세무대리인, 2;법인, 3;개인)
//			public char[] RPTGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//세무대리인
//			public char[] TAXAGE;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//홈텍스ID
//			public char[] HOMTID;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//세무프로그램코드
//			public char[] PGMCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] BUSNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//상      호
//			public char[] sangho;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//담당부서명
//			public char[] DAMDPT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//담당 성명
//			public char[] DAMNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(15), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 15)]
//				//담당전화번호
//			public char[] DAMTEL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//B Record수(신고의무자수)
//			public char[] BUSCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//한글코드종
//			public char[] HANCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(812), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 812)]
//				//공      란
//			public char[] FILLER;
//		}
//		A_record Arec;


//		private struct B_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] RECGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료  구분
//			public char[] DTAGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세  무  서
//			public char[] TAXCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련  번호
//			public char[] BUSCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] BUSNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//상      호
//			public char[] sangho;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//대  표  자
//			public char[] COMPRT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//C Record수
//			public char[] NEWCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//D Record수
//			public char[] OLDCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
//				//소득금액총액
//			public char[] INCOME;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//결정갑근세
//			public char[] GULGAB;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//    법인세
//			public char[] GULCOM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//    주민세
//			public char[] GULJUM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//    농특세
//			public char[] GULNON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//    총  액
//			public char[] GULTOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출대상기간
//			public char[] RNGCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(791), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 791)]
//				//공      란
//			public char[] FILLER;
//		}
//		B_record Brec;

///// 퇴직 주(현) 근무처 레코드 /
//		private struct C_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] RECGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료  구분
//			public char[] DTAGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세  무  서
//			public char[] TAXCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련  번호
//			public char[] SQNNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] BUSNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호(주현)
//			public char[] BUSNUM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//상      호
//			public char[] sangho;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//종전근무처수
//			public char[] JONCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//거주자구분
//			public char[] DWEGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//거주지국코드
//			public char[] RGNCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//귀속년도시작
//			public char[] STRINT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//귀속년도종료
//			public char[] ENDINT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//성      명
//			public char[] MSTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내외국인구분
//			public char[] INTGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//퇴직  사유
//			public char[] RETRES;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직 급여(법정)
//			public char[] TJKPAY;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직 급여(법정이외)-명예퇴직수당(추가퇴직금)
//			public char[] SUDAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직연금일시금(법정)(단체퇴직보험금)
//			public char[] BHMAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직연금일시금(법정 외)           (2010년 추가)
//			public char[] BHMAM1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직급여총액(법정)
//			public char[] TOTAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직급여총액(법정외 2011년추가)
//			public char[] TOTAM1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//비과세소득(2007년)
//			public char[] BTXPAY;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//퇴직연금 계좌번호                 (2010년 추가)
//			public char[] MYNACC;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금-총수령액(2007년)
//			public char[] MYNTOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금-월리금합계액(2007년)
//			public char[] MYNWON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금-소득자불입액(2007년)
//			public char[] MYNBUL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금-퇴직연금소득공제액(2007년)
//			public char[] MYNGON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금-퇴직연금일시금(2007년)
//			public char[] MYNYIL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//종(전) 퇴직연금 계좌번호          (2010년 추가)
//			public char[] JYNACC;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종(전) 퇴직연금-총수령액          (2007년 추가/2010년 위치이동)
//			public char[] JYNTOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종(전) 퇴직연금-원리금합계액      (2007년 추가/2010년 위치이동)
//			public char[] JYNWON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종(전) 퇴직연금-소득자불입액      (2007년 추가/2010년 위치이동)
//			public char[] JYNBUL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종(전) 퇴직연금-퇴직연금소득공제액(2007년 추가/2010년 위치이동)
//			public char[] JYNGON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종(전) 퇴직연금-퇴직연금일시금    (2007년 추가/2010년 위치이동)
//			public char[] JYITOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-지급예상액-주현              (2007년)
//			public char[] SH1JIG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-지급예상액-종전              (2009년)
//			public char[] SH3JIG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-과세이연금액-주현            (2010년 추가)
//			public char[] SH1IYO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-과세이연금액-종전            (2010년 추가)
//			public char[] SH3IYO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-기수령금액-주현              (2010년 추가)
//			public char[] SH1GIS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-기수령금액-종전              (2010년 추가)
//			public char[] SH3GIS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-총일시금                     (2007년)
//			public char[] SH1TIL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-수령가능퇴직급여액           (2007년)
//			public char[] SH1SUR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-환산퇴직소득공제             (2007년)
//			public char[] SH1GON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-환산퇴직소득과세표준         (2007년)
//			public char[] SH1GWA;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-환산연평균과세표준           (2007년)
//			public char[] SH1YAG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-연평균산출세액               (2007년)
//			public char[] SH1YAS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-지급예상액-주현       (2007년)
//			public char[] SH2JIG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-지급예상액-종전       (2009년)
//			public char[] SH4JIG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-과세이연금액-주현     (2010년 추가)
//			public char[] SH2IYO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-과세이연금액-종전     (2010년 추가)
//			public char[] SH4IYO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-기수령금액-주현       (2010년 추가)
//			public char[] SH2GIS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-기수령금액-종전       (2010년 추가)
//			public char[] SH4GIS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-총일시금              (2007년)
//			public char[] SH2TIL;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-수령가능퇴직급여액    (2007년)
//			public char[] SH2SUR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-환산퇴직소득공제      (2007년)
//			public char[] SH2GON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-환산퇴직소득과세표준  (2007년)
//			public char[] SH2GWA;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-환산연평균과세표준    (2007년)
//			public char[] SH2YAG;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//환산-법정외-연평균산출세액        (2007년)
//			public char[] SH2YAS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정-현 입사일자
//			public char[] INPDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정-현 퇴사일자
//			public char[] OUTDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정-현 근속월수
//			public char[] GNMMON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정-현 제외월수(2007년)
//			public char[] EXPMON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정-전 입사일자                  (2010년 위치이동)
//			public char[] JIPDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정-전 퇴사일자                  (2010년 위치이동)
//			public char[] JOTDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정-전 근속월수                  (2010년 위치이동)
//			public char[] JGMMON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정-전 제외월수(2007년)          (2010년 위치이동)
//			public char[] JEXMON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정-   중복월수
//			public char[] DUPMON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//법정-   근속연수
//			public char[] GNMYER;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정외-현  입사일자
//			public char[] FR2DAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정외-현  퇴사일자
//			public char[] TO2DAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정외-현  근속월수
//			public char[] GN2MON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정외-현  제외월수(2007년)
//			public char[] EX2MON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정외-전  입사일자               (2010년 위치이동)
//			public char[] JI2DAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//법정외-전  퇴사일자               (2010년 위치이동)
//			public char[] JO2DAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정외-전  근속월수               (2010년 위치이동)
//			public char[] JO2MON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정외-전  제외월수(2007년)       (2010년 위치이동)
//			public char[] JE2MON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//법정외-전  중복월수
//			public char[] JB2MON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//법정외-    근속연수
//			public char[] GN2YER;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-퇴직급여액
//			public char[] JS1RET;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-퇴직소득공제
//			public char[] JS1GON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-퇴직소득과세표준
//			public char[] JS1STD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-연평균과세표준
//			public char[] JS1YSD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-연평균산출세액
//			public char[] JS1YST;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-산출  세액
//			public char[] JS1SAN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정-세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
//			public char[] JS1FRN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-퇴직급여액
//			public char[] JS2RET;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-퇴직소득공제
//			public char[] JS2GON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-퇴직소득과세표준
//			public char[] JS2STD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-연평균과세표준
//			public char[] JS2YSD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-연평균산출세액
//			public char[] JS2YST;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//법정외-산출  세액
//			public char[] JS2SAN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
//			public char[] JS2FRN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-퇴직급여액
//			public char[] RETPAY;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-퇴직소득공제
//			public char[] RETGON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-퇴직소득과세표준
//			public char[] TAXSTD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-연평균과세표준
//			public char[] YTXSTD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-연평균산출세액
//			public char[] YSANTX;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-산출  세액
//			public char[] SANTAX;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//계-세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
//			public char[] TAXGON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//결정소득세
//			public char[] GULGAB;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    주민세
//			public char[] GULJUM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    특별세
//			public char[] GULNON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    세액계
//			public char[] GULTOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종전소득세
//			public char[] JONGAB;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    주민세
//			public char[] JONJUM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    특별세
//			public char[] JONNON;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//    세액계
//			public char[] JONTOT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//공      란
//			public char[] FILLER;
//		}
//		C_record Crec;

//		private struct D_Record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] RECGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료  구분
//			public char[] DTAGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세  무  서
//			public char[] TAXCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련  번호
//			public char[] SQNNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] BUSNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(50), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 50)]
//				//공      란
//			public char[] FILLD1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//근무처  명
//			public char[] JONNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] JONNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직  급여
//			public char[] RETAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//명예퇴직수당
//			public char[] SUDAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직연금일시금(법정)(퇴직보험금등)
//			public char[] BHMAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//퇴직연금일시금(법정외)
//			public char[] BHMAM1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종전계
//			public char[] TOTAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//종전계(법정이외)
//			public char[] TOTAM1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//비과세소득(2007년 추가)
//			public char[] BTXP01;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//종전근무처일련번호
//			public char[] JONCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(783), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 783)]
//				//공      란
//			public char[] FILLER;
//		}
//		D_Record Drec;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY421.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "ZPY421_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "ZPY421");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "Code"

//			oForm.Freeze(true);
//			ZPY421_CreateItems();
//			ZPY421_EnableMenus();
//			ZPY421_SetDocument(oFromDocEntry01);
//			//    Call ZPY421_FormResize

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

//		private bool ZPY421_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo1 = null;
//			SAPbouiCOM.ComboBox oCombo2 = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbouiCOM.OptionBtn optBtn = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 생성년도
//			oForm.DataSources.UserDataSources.Add("STRDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			oEdit = oForm.Items.Item("STRDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "STRDAT");

//			/// 생성년도
//			oForm.DataSources.UserDataSources.Add("ENDDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			oEdit = oForm.Items.Item("ENDDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "ENDDAT");

//			//Call oForm.DataSources.UserDataSources.Add("JsnGbn", dt_SHORT_TEXT, 10)     '/ 생성구분

//			/// 생성구분
//			oForm.DataSources.UserDataSources.Add("JSNTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo1 = oForm.Items.Item("JSNTYP").Specific;
//			oCombo1.ValidValues.Add("1", "연간(01.01~12.31)지급분");
//			oCombo1.ValidValues.Add("2", "폐업에 의한 수시 제출분");
//			oCombo1.ValidValues.Add("3", "수시 분할제출분");
//			oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);



//			///부서 From
//			oForm.DataSources.UserDataSources.Add("DptStr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 부서코드
//			oCombo1 = oForm.Items.Item("DptStr").Specific;

//			///부서 To
//			oForm.DataSources.UserDataSources.Add("DptEnd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo2 = oForm.Items.Item("DptEnd").Specific;

//			///사번
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");

//			///성명
//			oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");

//			///사번순번
//			oForm.DataSources.UserDataSources.Add("EmpID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oEdit = oForm.Items.Item("EmpID").Specific;
//			oEdit.DataBind.SetBound(true, "", "EmpID");

//			///제출일자
//			oForm.DataSources.UserDataSources.Add("PRTDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			oEdit = oForm.Items.Item("PRTDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "PRTDAT");

//			///경로
//			oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oEdit = oForm.Items.Item("Path").Specific;
//			oEdit.DataBind.SetBound(true, "", "Path");

//			//// 생성구분
//			//    Set oCombo1 = oForm.Items("JsnGbn").Specific
//			//    oCombo1.ValidValues.Add "1", "퇴직소득 지급명세서-1"
//			//    oCombo1.ValidValues.Add "2", "퇴직소득 지급명세서-2"
//			//    Call oCombo1.Select(0, psk_Index)   '/ 전체

//			//// 사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 지점
//			oCombo1 = oForm.Items.Item("CLTCOD").Specific;
//			oCombo1.DataBind.SetBound(true, "", "CLTCOD");

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			oForm.DataSources.UserDataSources.Add("Col0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oForm.DataSources.UserDataSources.Add("Col1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//			oColumn = oMat1.Columns.Item("Col0");
//			oColumn.DataBind.SetBound(true, "", "Col0");

//			oColumn = oMat1.Columns.Item("Col1");
//			oColumn.DataBind.SetBound(true, "", "Col1");

//			////----------------------------------------------------------------------------------------------
//			//// 데이터셋정의
//			////----------------------------------------------------------------------------------------------
//			//    '//테이블이 있을경우 데이터셋(Matrix)
//			//    Set oDS_ZPY421A = oForm.DataSources.DBDataSources("@ZPY421A")   '//헤더
//			//    Set oDS_ZPY421B = oForm.DataSources.DBDataSources("@ZPY421B")   '//라인
//			//
//			//    Set oMat1 = oForm.Items("Mat1").Specific       '
//			//
//			//    oMat1.SelectionMode = ms_NotSupported
//			//    oMat1.AutoResizeColumns

//			//    '//테이블이 없는경우 데이터셋(Grid)
//			//    oForm.DataSources.DataTables.Add ("PH_PY004")
//			//    oForm.DataSources.DataTables.Item("PH_PY004").Columns.Add "부서", ft_AlphaNumeric
//			//    oForm.DataSources.DataTables.Item("PH_PY004").Columns.Add "담당", ft_AlphaNumeric
//			//
//			//    Set oGrid1 = oForm.Items("Grid1").Specific
//			//
//			//    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY004")
//			//    Set oDS_PH_PY004 = oForm.DataSources.DataTables.Item("PH_PY004")


//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------
//			//    '//콤보1
//			//    Set oCombo1 = oForm.Items("    ").Specific
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo1)
//			//    oForm.Items("    ").DisplayDesc = True
//			//
//			//    '//콤보2
//			//    Set oCombo1 = oForm.Items("    ").Specific
//			//    oCombo1.ValidValues.Add "M", "남자"
//			//    oCombo1.ValidValues.Add "F", "여자"
//			//'    oCombo1.Select 0, psk_Index
//			//    oForm.Items("sex").DisplayDesc = True
//			//
//			//    '/체크박스
//			//    Set oCheck = oForm.Items("    ").Specific
//			//    oCheck.ValOn = "Y": oCheck.ValOff = "N"
//			//    oCheck.Checked = False
//			//
//			//    '//매트릭스컬럼
//			//    Set oColumn = oMat1.Columns("FILD01")
//			//    oColumn.Editable = True

//			//    '//UserDataSources
//			//    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
//			//    Set oCombo1 = oForm.Items("    ").Specific
//			//    oCombo1.DataBind.SetBound True, "", "    "
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo1)
//			//    oForm.Items("CLTCOD").DisplayDesc = True

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
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
//			ZPY421_CreateItems_Error:

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void ZPY421_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			ZPY421_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void ZPY421_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				ZPY421_FormItemEnabled();
//				ZPY421_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				ZPY421_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			ZPY421_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void ZPY421_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			int i = 0;
//			string sQry = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//
//				//        '//기본사항 - 부서 (사업장에 따른 부서변경)
//				//        Set oCombo = oForm.Items("DptStr").Specific
//				//
//				//        If oCombo.ValidValues.Count > 0 Then
//				//            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
//				//                oCombo.ValidValues.Remove i, psk_Index
//				//            Next i
//				//            oCombo.ValidValues.Add "", ""
//				//            oCombo.Select 0, psk_Index
//				//        End If
//				//
//				//        If oForm.Items("CLTCOD").Specific.Value <> "" Then
//				//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
//				//            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
//				//            sQry = sQry & " ORDER BY U_Code"
//				//            Call SetReDataCombo(oForm, sQry, oCombo)
//				//        End If

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//        '//기본사항 - 부서 (사업장에 따른 부서변경)
//				//        Set oCombo = oForm.Items("TeamCode").Specific
//				//
//				//        If oCombo.ValidValues.Count > 0 Then
//				//            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
//				//                oCombo.ValidValues.Remove i, psk_Index
//				//            Next i
//				//            oCombo.ValidValues.Add "", ""
//				//            oCombo.Select 0, psk_Index
//				//        End If
//				//
//				//        If oDS_PH_PY001A.GetValue("U_CLTCOD", 0) <> "" Then
//				//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
//				//            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
//				//            sQry = sQry & " ORDER BY U_Code"
//				//            Call SetReDataCombo(oForm, sQry, oCombo)
//				//        End If

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			ZPY421_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			ZPAY_g_EmpID MstInfo = default(ZPAY_g_EmpID);

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						/// ChooseBtn사원리스트
//						if (pval.ItemUID == "CBtn1") {

//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else if (pval.ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//							if (ZPY421_DataValidCheck() == false) {
//								BubbleEvent = false;
//								return;
//							}
//							if (File_Create() == false) {
//								BubbleEvent = false;
//								return;
//							} else {
//								BubbleEvent = false;
//								oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//							}
//						} else if (pval.ItemUID == "Btn1") {
//							oFilePath = My.MyProject.Forms.ZP_Form.vbGetBrowseDirectory(ref ZP_Form);
//							oForm.DataSources.UserDataSources.Item("Path").ValueEx = oFilePath;
//							BubbleEvent = false;
//							return;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "") == true) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat1":
//						case "Grid1":
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
//							if (pval.ItemUID == "CLTCOD") {

//								////기본사항 - 부서 (사업장에 따른 부서변경)
//								oCombo = oForm.Items.Item("DptStr").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//									oCombo.ValidValues.Add("", "");
//									oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//								sQry = sQry + " ORDER BY U_Code";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//								oCombo.ValidValues.Add("%", "전체");
//								oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//								oForm.Items.Item("DptStr").DisplayDesc = true;

//								////기본사항 - 부서 (사업장에 따른 부서변경)
//								oCombo = oForm.Items.Item("DptEnd").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//									oCombo.ValidValues.Add("", "");
//									oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//								sQry = sQry + " ORDER BY U_Code";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//								oCombo.ValidValues.Add("%", "전체");
//								oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//								oForm.Items.Item("DptEnd").DisplayDesc = true;


//							}
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						///정산년도
//						if (oLastItemUID == "STRDAT") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oLastItemUID).Specific.VALUE))) {
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (MDC_SetMod.ChkYearMonth(ref Strings.Trim(Convert.ToString(oForm.Items.Item(oLastItemUID).Specific.VALUE))) == false) {
//									oForm.Items.Item(oLastItemUID).Update();
//									MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//								}
//							}
//						} else if (oLastItemUID == "MSTCOD") {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(oLastItemUID).Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item(oLastItemUID).Specific.String) + "'", ref "") == true) {
//								oForm.Items.Item(oLastItemUID).Update();
//								MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//								BubbleEvent = false;
//							}
//						}
//					}
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
//						if (pval.ItemUID == "STRDAT") {
//							//oJsnYear = Mid$(oForm.Items(oUID).Specific.String, 1, 4)
//							oJsnYear = Strings.Left(oForm.DataSources.UserDataSources.Item("STRDAT").ValueEx, 4);
//							oForm.DataSources.UserDataSources.Item("ENDDAT").ValueEx = oJsnYear + "1231";
//							oForm.Items.Item("ENDDAT").Update();
//						}

//						if (pval.ItemUID == "ENDDAT") {

//						}

//						if (pval.ItemUID == "MSTCOD") {
//							//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.String)) {
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oForm.Items.Item("MSTCOD").Specific.String = "";
//								oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//								oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = "";
//							} else {
//								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oForm.Items.Item("MSTCOD").Specific.String = Strings.UCase(oForm.Items.Item("MSTCOD").Specific.String);
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								MstInfo = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item("MSTCOD").Specific.String);
//								oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MstInfo.MSTNAM;
//								oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = MstInfo.EmpID;
//							}
//							oForm.Items.Item("MSTNAM").Update();
//							oForm.Items.Item("EmpID").Update();
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
//				//                Call ZPY421_FormItemEnabled
//				//                Call ZPY421_AddMatrixRow
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
//						//                Set oDS_ZPY421A = Nothing
//						//                Set oDS_ZPY421B = Nothing
//						//
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
//						ZPY421_FormItemEnabled();
//						ZPY421_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						ZPY421_FormItemEnabled();
//						ZPY421_AddMatrixRow();
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						ZPY421_FormItemEnabled();
//						ZPY421_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						ZPY421_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;
//					//                '// [MAT1 용]
//					//                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//					//                    oMat1.FlushToDataSource
//					//
//					//                    While (i <= oDS_ZPY421B.Size - 1)
//					//                        If oDS_ZPY421B.GetValue("U_FILD01", i) = "" Then
//					//                            oDS_ZPY421B.RemoveRecord (i)
//					//                            i = 0
//					//                        Else
//					//                            i = i + 1
//					//                        End If
//					//                    Wend
//					//
//					//                    For i = 0 To oDS_ZPY421B.Size
//					//                        Call oDS_ZPY421B.setValue("U_LineNum", i, i + 1)
//					//                    Next i
//					//
//					//                    oMat1.LoadFromDataSource
//					//                End If
//					//                Call ZPY421_AddMatrixRow
//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

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

//		public void ZPY421_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			//    '//[Mat1 용]
//			//    oMat1.FlushToDataSource
//			//    oRow = oMat1.VisualRowCount
//			//
//			//    If oMat1.VisualRowCount > 0 Then
//			//        If Trim(oDS_ZPY421B.GetValue("U_FILD01", oRow - 1)) <> "" Then
//			//            If oDS_ZPY421B.Size <= oMat1.VisualRowCount Then
//			//                oDS_ZPY421B.InsertRecord (oRow)
//			//            End If
//			//            oDS_ZPY421B.Offset = oRow
//			//            oDS_ZPY421B.setValue "U_LineNum", oRow, oRow + 1
//			//            oDS_ZPY421B.setValue "U_FILD01", oRow, ""
//			//            oDS_ZPY421B.setValue "U_FILD02", oRow, ""
//			//            oDS_ZPY421B.setValue "U_FILD03", oRow, 0
//			//            oMat1.LoadFromDataSource
//			//        Else
//			//            oDS_ZPY421B.Offset = oRow - 1
//			//            oDS_ZPY421B.setValue "U_LineNum", oRow - 1, oRow
//			//            oDS_ZPY421B.setValue "U_FILD01", oRow - 1, ""
//			//            oDS_ZPY421B.setValue "U_FILD02", oRow - 1, ""
//			//            oDS_ZPY421B.setValue "U_FILD03", oRow - 1, 0
//			//            oMat1.LoadFromDataSource
//			//        End If
//			//    ElseIf oMat1.VisualRowCount = 0 Then
//			//        oDS_ZPY421B.Offset = oRow
//			//        oDS_ZPY421B.setValue "U_LineNum", oRow, oRow + 1
//			//        oDS_ZPY421B.setValue "U_FILD01", oRow, ""
//			//        oDS_ZPY421B.setValue "U_FILD02", oRow, ""
//			//        oDS_ZPY421B.setValue "U_FILD03", oRow, 0
//			//        oMat1.LoadFromDataSource
//			//    End If

//			oForm.Freeze(false);
//			return;
//			ZPY421_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void ZPY421_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'ZPY421'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			ZPY421_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool ZPY421_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			functionReturnValue = false;

//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("STRDAT").Specific.String)) | string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ENDDAT").Specific.String))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("대상일자를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				return functionReturnValue;
//				//UPGRADE_WARNING: oForm.Items(JSNTYP).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("JSNTYP").Specific.Selected == null) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("기간 코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				return functionReturnValue;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("PRTDAT").Specific.String))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				return functionReturnValue;
//				//UPGRADE_WARNING: oForm.Items(STRDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oForm.Items(ENDDAT).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("ENDDAT").Specific.String < oForm.Items.Item("STRDAT").Specific.String) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("대상기간종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				return functionReturnValue;
//			}

//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTSTR = oForm.Items.Item("DptStr").Specific.Selected.VALUE;
//			if (DPTSTR == "%")
//				DPTSTR = "00000001";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTEND = oForm.Items.Item("DptEnd").Specific.Selected.VALUE;
//			if (DPTEND == "%")
//				DPTEND = "ZZZZZZZZ";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
//				MSTCOD = "%";
//			//    oJsnGbn = oForm.Items("JsnGbn").Specific.Selected.Value

//			////----------------------------------------------------------------------------------
//			////필수 체크
//			////----------------------------------------------------------------------------------

//			//    oMat1.FlushToDataSource
//			//    '// Matrix 마지막 행 삭제(DB 저장시)
//			//    If oDS_ZPY421B.Size > 1 Then oDS_ZPY421B.RemoveRecord (oDS_ZPY421B.Size - 1)
//			//    oMat1.LoadFromDataSource

//			functionReturnValue = true;

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			ZPY421_DataValidCheck_Error:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool ZPY421_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [ZPY421A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@ZPY421A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto ZPY421_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			ZPY421_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			ZPY421_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}



//		private void ZPY421_Print_Report01()
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
//				goto ZPY421_Print_Report01_Error;
//			}

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//			WinTitle = "[S142] 발주서";
//			ReportName = "S142_1.rpt";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "EXEC ZPY421_1 '" + oForm.Items.Item("8").Specific.VALUE + "'";
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
//			ZPY421_Print_Report01_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("ZPY421_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}



//		private bool File_Create()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string oStr = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset sRecordset = null;

//			sRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			ErrNum = 0;
//			/// Question
//			if (MDC_Globals.Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2) {
//				ErrNum = 1;
//				goto Error_Message;
//			}
//			oMat1.Clear();
//			MaxRow = 0;
//			BUSCNT = 0;
//			/// B레크드 일련번호
//			BUSTOT = 0;

//			/// 파일경로설정
//			if (string.IsNullOrEmpty(oFilePath))
//				oFilePath = "C:\\EOSDATA";
//			oFilePath = (Strings.Right(oFilePath, 1) == "\\" ? oFilePath : oFilePath + "\\");
//			oStr = MDC_SetMod.CreateFolder(ref Strings.Trim(oFilePath));
//			if (!string.IsNullOrEmpty(Strings.Trim(oStr))) {
//				ErrNum = 5;
//				goto Error_Message;
//			}

//			/// 갑근 제출자(대리인) 레코드
//			if (File_Create_ARecord() == false) {
//				ErrNum = 2;
//				goto Error_Message;
//			}

//			FileSystem.FileClose(1);
//			FileSystem.FileOpen(1, FILNAM.Value, OpenMode.Output);
//			/// A레코드: 갑근 원천징수의무자별 집계 레코드
//			FileSystem.PrintLine(1, MDC_SetMod.sStr(ref Arec.RECGBN) + MDC_SetMod.sStr(ref Arec.DTAGBN) + MDC_SetMod.sStr(ref Arec.TAXCOD) + MDC_SetMod.sStr(ref Arec.PRTDAT) + MDC_SetMod.sStr(ref Arec.RPTGBN) + MDC_SetMod.sStr(ref Arec.TAXAGE) + MDC_SetMod.sStr(ref Arec.HOMTID) + MDC_SetMod.sStr(ref Arec.PGMCOD) + MDC_SetMod.sStr(ref Arec.BUSNBR) + MDC_SetMod.sStr(ref Arec.sangho) + MDC_SetMod.sStr(ref Arec.DAMDPT) + MDC_SetMod.sStr(ref Arec.DAMNAM) + MDC_SetMod.sStr(ref Arec.DAMTEL) + MDC_SetMod.sStr(ref Arec.BUSCNT) + MDC_SetMod.sStr(ref Arec.HANCOD) + MDC_SetMod.sStr(ref Arec.FILLER));

//			Matrix_AddRow("제출자 레코드 생성 완료!", true);
//			/// B레코드: 갑근 집계 레코드 /***********************************************/
//			sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] T0 ";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " WHERE U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.VALUE + "' ORDER BY CODE";
//			sRecordset.DoQuery(sQry);
//			while (!(sRecordset.EoF)) {
//				NEWCNT = 0;
//				//UPGRADE_WARNING: sRecordset.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CLTCOD = sRecordset.Fields.Item(0).Value;

//				//UPGRADE_WARNING: sRecordset.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Brec.TAXCOD = sRecordset.Fields.Item("U_TAXCODE").Value;
//				Brec.BUSNBR = Strings.Replace(sRecordset.Fields.Item("U_BUSNUM").Value, "-", "");
//				Brec.sangho = Strings.Trim(sRecordset.Fields.Item("U_CLTNAME").Value);
//				Brec.COMPRT = Strings.Trim(sRecordset.Fields.Item("U_COMPRT").Value);
//				Brec.PERNBR = Strings.Replace(Strings.Trim(sRecordset.Fields.Item("U_PERNUM").Value), "-", "");

//				/// B레코드: 갑근 집계 레코드
//				switch (File_Create_BRecord()) {
//					case 0:
//						Matrix_AddRow(CLTCOD + "- 징수의무자의 집계 레코드 생성 완료!", true);

//						/// C레코드: 갑근 주(현)근무처 레코드
//						if (File_Create_CRecord() == false) {
//							ErrNum = 4;
//							goto Error_Message;
//						}
//						Matrix_AddRow(CLTCOD + "- 징수의무자의 데이터 레코드" + NEWCNT + "건 생성 완료!", true);
//						break;
//					case 1:
//						//// B레코드가 생성되지 않은 경우 C레코드 생성하지 않고, 건너뜀
//						break;

//					case 2:
//						ErrNum = 3;
//						goto Error_Message;
//						break;
//				}

//				sRecordset.MoveNext();
//			}
//			/// A레코드의 원천의무자수/B레코드의 원천의무자수
//			if (BUSTOT != BUSCNT) {
//				ErrNum = 6;
//				goto Error_Message;
//			}
//			FileSystem.FileClose(1);
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;

//			oForm.DataSources.UserDataSources.Item("Path").Value = FILNAM.Value;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			functionReturnValue = true;
//			return functionReturnValue;
//			Error_Message:

//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("A레코드(갑근 제출자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B레코드(갑근 원천징수의무자별 집계 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드(갑근 주(현)근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("CreateFolder Error : " + oStr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 6) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B레코드의 징수의무자수(" + Convert.ToString(BUSCNT) + ")와 A레코드의 신고의무자수(" + Convert.ToString(BUSTOT) + ")가 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private bool File_Create_ARecord()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string PRTDAT = null;
//			string BUSNUM = null;
//			string CheckA = null;

//			CheckA = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PRTDAT = Strings.Mid(oForm.Items.Item("PRTDAT").Specific.String, 1, 4) + Strings.Mid(oForm.Items.Item("PRTDAT").Specific.String, 6, 2) + Strings.Mid(oForm.Items.Item("PRTDAT").Specific.String, 9, 2);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 신고의무자수

//			sQry = " SELECT  COUNT(DISTINCT ISNULL(T1.U_CLTCOD, ''))";
//			sQry = sQry + " FROM [@PH_PY115A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code";
//			sQry = sQry + " INNER JOIN [@PS_HR200L] T2 ON T1.U_TeamCode = T2.U_Code AND T2.Code = '1'";
//			sQry = sQry + " INNER JOIN [@PH_PY005A] T3 ON ISNULL(T1.U_CLTCOD, '') = T3.Code";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " WHERE   T3.U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.VALUE + "'";
//			sQry = sQry + " AND     T2.U_Code BETWEEN " + "'" + DPTSTR + "'" + " AND " + "'" + DPTEND + "'";
//			sQry = sQry + " AND     T0.U_MSTCOD LIKE N'%'";
//			sQry = sQry + " AND     T0.U_RETPAY > 0";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND     T0.U_ENDINT >= '" + oForm.Items.Item("STRDAT").Specific.VALUE + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND     T0.U_ENDINT <= '" + oForm.Items.Item("ENDDAT").Specific.VALUE + "'";


//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				BUSTOT = oRecordSet.Fields.Item(0).Value;
//			}
//			if (Conversion.Val(Convert.ToString(BUSTOT)) == 0) {
//				ErrNum = 3;
//				goto Error_Message;
//			}

//			/// 업체정보
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" + oForm.Items.Item("CLTCOD").Specific.Selected.VALUE + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				/// 파일명
//				BUSNUM = Strings.Replace(oRecordSet.Fields.Item("U_BUSNUM").Value, "-", "");
//				if (Strings.Len(Strings.Trim(BUSNUM)) != 10) {
//					ErrNum = 2;
//					goto Error_Message;
//				}
//				//FILNAM = "C:\EOSDATA\E" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
//				//FILNAM = oFilePath & "E" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
//				FILNAM.Value = oFilePath + "EA" + Strings.Mid(BUSNUM, 1, 7) + "." + Strings.Mid(BUSNUM, 8, 3);
//				/// 2010년 변경

//				/// B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				BUSNUM = Strings.Replace(oRecordSet.Fields.Item("U_BUSNUM").Value, "-", "");
//				/// A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				Arec.RECGBN = "A";
//				/// 레코드 구분
//				//Arec.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
//				Arec.DTAGBN = "25";
//				/// 자료구분
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Arec.TAXCOD = oRecordSet.Fields.Item("U_TAXCODE").Value;
//				/// 세무서
//				Arec.PRTDAT = PRTDAT;
//				/// 제출일자
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Arec.RPTGBN = oRecordSet.Fields.Item("U_TaxDGbn").Value;
//				/// 제출자
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Arec.TAXAGE = oRecordSet.Fields.Item("U_TaxDCode").Value;
//				/// 세무대리인
//				Arec.HOMTID = Strings.Trim(oRecordSet.Fields.Item("U_HOMETID").Value);
//				/// 홈텍스ID
//				Arec.PGMCOD = "9000";
//				/// 세무프로그램코드
//				Arec.BUSNBR = Strings.Replace(oRecordSet.Fields.Item("U_TAXDBUS").Value, "-", "");
//				/// 사업자번호
//				Arec.sangho = Strings.Trim(oRecordSet.Fields.Item("U_TAXDNAM").Value);
//				/// 상호
//				Arec.DAMDPT = Strings.Trim(oRecordSet.Fields.Item("U_CHGDPT").Value);
//				/// 담당부서명
//				Arec.DAMNAM = Strings.Trim(oRecordSet.Fields.Item("U_CHGNAME").Value);
//				/// 담당성명
//				Arec.DAMTEL = Strings.Trim(oRecordSet.Fields.Item("U_CHGTEL").Value);
//				/// 담당전화번호
//				Arec.BUSCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(BUSTOT, new string("0", Strings.Len(Arec.BUSCNT)));
//				/// 원천징수의무자수
//				Arec.HANCOD = "101";
//				/// 한글코드 종류
//				Arec.FILLER = Strings.Space(Strings.Len(Arec.FILLER));

//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.TAXCOD))) {
//					Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.RPTGBN))) {
//					Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.BUSNBR))) {
//					Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.sangho))) {
//					Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.DAMDPT))) {
//					Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.DAMNAM))) {
//					Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Arec.DAMTEL))) {
//					Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//			}

//			if (Convert.ToBoolean(CheckA) == false) {
//				functionReturnValue = true;
//			} else {
//				functionReturnValue = false;
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			Error_Message:
//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도의 자사정보가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사정보등록의 사업자번호가 올바르지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("자사정보등록의 신고의무사업장이 존재하지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				Matrix_AddRow("A레코드오류: " + Err().Description, ref false, ref true);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private short File_Create_BRecord()
//		{
//			short functionReturnValue = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckB = null;

//			CheckB = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 집계정보
//			sQry = " SELECT   ISNULL(COUNT(T0.U_MSTCOD),0) AS SUM_NEWCNT, ";
//			sQry = sQry + "   SUM(CASE WHEN ISNULL(T0.U_J01NAM,'')='' THEN 0 ELSE 1 END) AS SUM_OLDCNT, ";
//			sQry = sQry + "   ISNULL(SUM(T0.U_RETPAY),0)  AS SUM_INCOME, ";
//			sQry = sQry + "   ISNULL(SUM(T0.U_GULGAB),0)  AS SUM_GULGAB, ";
//			sQry = sQry + "   ISNULL(SUM(T0.U_GULJUM),0)  AS SUM_GULJUM, ";
//			sQry = sQry + "   0                           AS SUM_GULNON  ";
//			sQry = sQry + "  FROM [@PH_PY115A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code ";
//			sQry = sQry + "                        INNER JOIN [OUDP] T2 ON T1.Dept = T2.Code                                  ";
//			sQry = sQry + "  WHERE   ISNULL(T1.U_CLTCOD, '') = '" + Strings.Trim(CLTCOD) + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + "  AND     ISNULL(T1.Branch, '') LIKE " + "N'" + oForm.Items.Item("BPLId").Specific.Selected.VALUE + "'";
//			sQry = sQry + "  AND     T1.U_TeamCode BETWEEN " + "'" + DPTSTR + "'" + " AND " + "'" + DPTEND + "'";
//			sQry = sQry + "  AND     T0.U_MSTCOD LIKE " + "N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + "  AND     T0.U_RETPAY > 0";
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + "  AND     T0.U_ENDINT >= '" + oForm.Items.Item("STRDAT").Specific.String + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + "  AND     T0.U_ENDINT <= '" + oForm.Items.Item("ENDDAT").Specific.String + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else if (oRecordSet.Fields.Item(0).Value == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				BUSCNT = BUSCNT + 1;
//				/// B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				Brec.RECGBN = "B";
//				/// 레코드구분
//				//Brec.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
//				Brec.DTAGBN = "25";
//				/// 자료구분
//				Brec.TAXCOD = Arec.TAXCOD;
//				/// 세무서
//				Brec.BUSCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(BUSCNT, new string("0", Strings.Len(Brec.BUSCNT)));
//				/// 원천징수의무자수 일련번호

//				Brec.NEWCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_NEWCNT").Value, new string("0", Strings.Len(Brec.NEWCNT)));
//				//C Record수
//				Brec.OLDCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_OLDCNT").Value, new string("0", Strings.Len(Brec.OLDCNT)));
//				//D Record수
//				Brec.INCOME = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_INCOME").Value, new string("0", Strings.Len(Brec.INCOME)));
//				//소득금액 총액
//				Brec.GULGAB = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_GULGAB").Value, new string("0", Strings.Len(Brec.GULGAB)));
//				//결정 소득세
//				Brec.GULCOM = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Brec.GULCOM)));
//				//공란
//				Brec.GULJUM = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_GULJUM").Value, new string("0", Strings.Len(Brec.GULJUM)));
//				//결정 주민세
//				Brec.GULNON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_GULNON").Value, new string("0", Strings.Len(Brec.GULNON)));
//				//결정 농특세
//				Brec.GULTOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("SUM_GULGAB").Value + oRecordSet.Fields.Item("SUM_GULJUM").Value + oRecordSet.Fields.Item("SUM_GULNON").Value, new string("0", Strings.Len(Brec.GULTOT)));
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Brec.RNGCOD = oForm.Items.Item("JSNTYP").Specific.Selected.VALUE;
//				//제출대상 기간
//				Brec.FILLER = Strings.Space(Strings.Len(Brec.FILLER));

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref Brec.RECGBN) + MDC_SetMod.sStr(ref Brec.DTAGBN) + MDC_SetMod.sStr(ref Brec.TAXCOD) + MDC_SetMod.sStr(ref Brec.BUSCNT) + MDC_SetMod.sStr(ref Brec.BUSNBR) + MDC_SetMod.sStr(ref Brec.sangho) + MDC_SetMod.sStr(ref Brec.COMPRT) + MDC_SetMod.sStr(ref Brec.PERNBR) + MDC_SetMod.sStr(ref Brec.NEWCNT) + MDC_SetMod.sStr(ref Brec.OLDCNT) + MDC_SetMod.sStr(ref Brec.INCOME) + MDC_SetMod.sStr(ref Brec.GULGAB) + MDC_SetMod.sStr(ref Brec.GULCOM) + MDC_SetMod.sStr(ref Brec.GULJUM) + MDC_SetMod.sStr(ref Brec.GULNON) + MDC_SetMod.sStr(ref Brec.GULTOT) + MDC_SetMod.sStr(ref Brec.RNGCOD) + MDC_SetMod.sStr(ref Brec.FILLER));


//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(Brec.BUSNBR))) {
//					Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckB = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Brec.COMPRT))) {
//					Matrix_AddRow("B레코드:대표자명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckB = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(Brec.PERNBR))) {
//					Matrix_AddRow("B레코드:법인(주민)등록번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckB = Convert.ToString(true);
//				}
//			}

//			if (Convert.ToBoolean(CheckB) == false) {
//				functionReturnValue = 0;
//			} else {
//				functionReturnValue = 2;
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			Error_Message:
//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("집계레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				functionReturnValue = 1;
//			} else {
//				Matrix_AddRow("B레코드오류: " + Err().Description, false);
//				functionReturnValue = 2;
//			}
//			return functionReturnValue;

//		}

//		private bool File_Create_CRecord()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckC = null;
//			CheckC = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;
//			C_MSTCOD = "";
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 사원별 정산정보
//			sQry = " SELECT  ISNULL(T4.U_DWEGBN, 0)   AS U_DWEGBN,";
//			sQry = sQry + "  ISNULL(T1.brthcountr,'') AS U_DWECOD,";
//			sQry = sQry + "  ISNULL(T1.citizenshp,'') AS U_INTCOD,";
//			sQry = sQry + "  ISNULL(T4.U_FRGTAX,2)   AS U_FRGTAX,";
//			sQry = sQry + "  ISNULL(T4.U_INTGBN,1)   AS U_INTGBN,";
//			sQry = sQry + "  ISNULL(T1.govID,'')     AS U_PERNBR,";
//			sQry = sQry + "  CASE WHEN T0.U_ST2RET IS NULL THEN T0.U_JINDAT ELSE DATEADD(DD,1,T0.U_ST2RET) END AS ST2RET,";
//			sQry = sQry + "  T0.*";
//			sQry = sQry + " FROM [@PH_PY115A] T0 ";
//			sQry = sQry + "      INNER JOIN [OHEM] T1 ON T0.U_EmpID = T1.EmpID";
//			sQry = sQry + "      INNER JOIN [@PH_PY001A] T4 ON T0.U_MSTCOD = T4.Code";
//			sQry = sQry + " WHERE   ISNULL(T1.U_CLTCOD, '') = '" + Strings.Trim(CLTCOD) + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND     ISNULL(T1.Branch, '') LIKE " + "N'" + oForm.Items.Item("BPLId").Specific.Selected.VALUE + "'";
//			sQry = sQry + " AND     T4.U_TeamCode BETWEEN " + "'" + DPTSTR + "'" + " And " + "'" + DPTEND + "'";
//			sQry = sQry + " AND     T0.U_MSTCOD LIKE " + "N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + " AND     T0.U_RETPAY > 0";
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND     T0.U_ENDINT >= '" + oForm.Items.Item("STRDAT").Specific.String + "'";
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND     T0.U_ENDINT <= '" + oForm.Items.Item("ENDDAT").Specific.String + "'";
//			sQry = sQry + " ORDER BY T0.DocEntry";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				//        ErrNum = 1
//				//        GoTo error_Message
//			}
//			while (!(oRecordSet.EoF)) {
//				NEWCNT = NEWCNT + 1;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_MSTCOD = oRecordSet.Fields.Item("U_MSTCOD").Value;
//				/// C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				Crec.RECGBN = "C";
//				Crec.DTAGBN = "25";
//				Crec.TAXCOD = Arec.TAXCOD;
//				Crec.SQNNBR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(Crec.SQNNBR)));
//				/// 일련번호
//				Crec.BUSNBR = Brec.BUSNBR;
//				Crec.BUSNUM = Brec.BUSNBR;
//				Crec.sangho = Brec.sangho;

//				if (!string.IsNullOrEmpty(Strings.Trim(oRecordSet.Fields.Item("U_J01NAM").Value))) {
//					Crec.JONCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(1, new string("0", Strings.Len(Crec.JONCNT)));
//					/// 종전근무처수
//				} else {
//					Crec.JONCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.JONCNT)));
//					/// 종전근무처수
//				}
//				Crec.DWEGBN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_DWEGBN").Value, new string("0", Strings.Len(Crec.DWEGBN)));
//				Crec.INTGBN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_INTGBN").Value, new string("0", Strings.Len(Crec.INTGBN)));
//				if (Crec.DWEGBN != "1") {
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					Crec.RGNCOD = oRecordSet.Fields.Item("U_DWECOD").Value;
//					/// 거주지국
//				} else if (Crec.INTGBN != "1") {
//					//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					Crec.RGNCOD = oRecordSet.Fields.Item("U_INTCOD").Value;
//					/// 거주지국
//				} else {
//					//If Crec.DWEGBN = "1" And Crec.INTGBN = "1" Then '/ 비거주자인경우 거주지국코드
//					Crec.RGNCOD = "";
//				}
//				Crec.STRINT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(Strings.Replace(oRecordSet.Fields.Item("U_STRINT").Value, "-", ""), 1, 8), new string("0", Strings.Len(Crec.STRINT)));
//				Crec.ENDINT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(Strings.Replace(oRecordSet.Fields.Item("U_ENDINT").Value, "-", ""), 1, 8), new string("0", Strings.Len(Crec.ENDINT)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Crec.MSTNAM = oRecordSet.Fields.Item("U_MSTNAM").Value;
//				Crec.PERNBR = Strings.Replace(oRecordSet.Fields.Item("U_PERNBR").Value, "-", "");
//				Crec.RETRES = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_RETRES").Value, new string("0", Strings.Len(Crec.RETRES)));
//				//퇴직  사유
//				Crec.TJKPAY = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_TJKPAY").Value, new string("0", Strings.Len(Crec.TJKPAY)));
//				Crec.SUDAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SUDAMT").Value, new string("0", Strings.Len(Crec.SUDAMT)));
//				Crec.BHMAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_YILPA1").Value, new string("0", Strings.Len(Crec.BHMAMT)));
//				/// 퇴직연금일시금
//				Crec.BHMAM1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_YILPA2").Value, new string("0", Strings.Len(Crec.BHMAM1)));
//				/// 컬럼을 안만들어놔서 일단 0으로 처리
//				Crec.TOTAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_TJKPAY").Value + oRecordSet.Fields.Item("U_YILPA1").Value, new string("0", Strings.Len(Crec.TOTAMT)));
//				Crec.TOTAM1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SUDAMT").Value + oRecordSet.Fields.Item("U_YILPA2").Value, new string("0", Strings.Len(Crec.TOTAM1)));
//				Crec.BTXPAY = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_BTXPAY").Value, new string("0", Strings.Len(Crec.BTXPAY)));

//				Crec.MYNACC = Strings.Trim(oRecordSet.Fields.Item("U_MYNACC").Value);
//				Crec.MYNTOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MYNTOT").Value, new string("0", Strings.Len(Crec.MYNTOT)));
//				Crec.MYNWON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MYNWON").Value, new string("0", Strings.Len(Crec.MYNWON)));
//				Crec.MYNBUL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MYNBUL").Value, new string("0", Strings.Len(Crec.MYNBUL)));
//				Crec.MYNGON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MYNGON").Value, new string("0", Strings.Len(Crec.MYNGON)));
//				Crec.MYNYIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_YILPAY").Value, new string("0", Strings.Len(Crec.MYNYIL)));

//				Crec.JYNACC = Strings.Trim(oRecordSet.Fields.Item("U_JYNACC").Value);
//				Crec.JYNTOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JYNTOT").Value, new string("0", Strings.Len(Crec.JYNTOT)));
//				Crec.JYNWON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JYNWON").Value, new string("0", Strings.Len(Crec.JYNWON)));
//				Crec.JYNBUL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JYNBUL").Value, new string("0", Strings.Len(Crec.JYNBUL)));
//				Crec.JYNGON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JYNGON").Value, new string("0", Strings.Len(Crec.JYNGON)));
//				Crec.JYITOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JYIL01").Value + oRecordSet.Fields.Item("U_JYIL02").Value, new string("0", Strings.Len(Crec.JYITOT)));

//				Crec.SH1JIG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1JIG").Value, new string("0", Strings.Len(Crec.SH1JIG)));
//				Crec.SH3JIG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH3JIG").Value, new string("0", Strings.Len(Crec.SH3JIG)));
//				Crec.SH1IYO = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1IYO").Value, new string("0", Strings.Len(Crec.SH1IYO)));
//				//// 컬럼만 만들어놓고 화면에는 아직 없어용
//				Crec.SH3IYO = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH3IYO").Value, new string("0", Strings.Len(Crec.SH3IYO)));
//				//// 컬럼만 만들어놓고 화면에는 아직 없어용
//				Crec.SH1GIS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1GIS)));
//				//// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
//				Crec.SH3GIS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH3GIS)));
//				//// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
//				if (Conversion.Val(Crec.SH1JIG) + Conversion.Val(Crec.SH3JIG) > 0) {
//					Crec.SH1TIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1TIL").Value, new string("0", Strings.Len(Crec.SH1TIL)));
//					Crec.SH1SUR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1SUR").Value, new string("0", Strings.Len(Crec.SH1SUR)));
//					Crec.SH1GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1GON").Value, new string("0", Strings.Len(Crec.SH1GON)));
//					Crec.SH1GWA = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1GWA").Value, new string("0", Strings.Len(Crec.SH1GWA)));
//					Crec.SH1YAG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1YAG").Value, new string("0", Strings.Len(Crec.SH1YAG)));
//					Crec.SH1YAS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH1YAS").Value, new string("0", Strings.Len(Crec.SH1YAS)));
//				} else {
//					Crec.SH1TIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1TIL)));
//					Crec.SH1SUR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1SUR)));
//					Crec.SH1GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1GON)));
//					Crec.SH1GWA = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1GWA)));
//					Crec.SH1YAG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1YAG)));
//					Crec.SH1YAS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH1YAS)));
//				}

//				Crec.SH2JIG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2JIG").Value, new string("0", Strings.Len(Crec.SH2JIG)));
//				Crec.SH4JIG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH4JIG").Value, new string("0", Strings.Len(Crec.SH4JIG)));
//				Crec.SH2IYO = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2IYO").Value, new string("0", Strings.Len(Crec.SH2IYO)));
//				//// 컬럼만 만들어놓고 화면에는 아직 없어용
//				Crec.SH4IYO = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH4IYO").Value, new string("0", Strings.Len(Crec.SH4IYO)));
//				//// 컬럼만 만들어놓고 화면에는 아직 없어용
//				Crec.SH2GIS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2GIS)));
//				//// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
//				Crec.SH4GIS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH4GIS)));
//				//// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
//				if (Conversion.Val(Crec.SH2JIG) + Conversion.Val(Crec.SH4JIG) > 0) {
//					Crec.SH2TIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2TIL").Value, new string("0", Strings.Len(Crec.SH2TIL)));
//					Crec.SH2SUR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2SUR").Value, new string("0", Strings.Len(Crec.SH2SUR)));
//					Crec.SH2GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2GON").Value, new string("0", Strings.Len(Crec.SH2GON)));
//					Crec.SH2GWA = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2GWA").Value, new string("0", Strings.Len(Crec.SH2GWA)));
//					Crec.SH2YAG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2YAG").Value, new string("0", Strings.Len(Crec.SH2YAG)));
//					Crec.SH2YAS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2YAS").Value, new string("0", Strings.Len(Crec.SH2YAS)));
//				} else {
//					Crec.SH2TIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2TIL)));
//					Crec.SH2SUR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2SUR)));
//					Crec.SH2GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2GON)));
//					Crec.SH2GWA = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2GWA)));
//					Crec.SH2YAG = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2YAG)));
//					Crec.SH2YAS = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.SH2YAS)));
//				}

//				/// 법정 주현
//				Crec.INPDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_STRRET").Value, "-", ""), new string("0", Strings.Len(Crec.INPDAT)));
//				Crec.OUTDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_ENDRET").Value, "-", ""), new string("0", Strings.Len(Crec.OUTDAT)));
//				Crec.GNMMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMMON").Value, new string("0", Strings.Len(Crec.GNMMON)));
//				Crec.EXPMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_EXPMON").Value, new string("0", Strings.Len(Crec.EXPMON)));
//				/// 법정 종전
//				Crec.JIPDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_JINDAT").Value, "-", ""), new string("0", Strings.Len(Crec.JIPDAT)));
//				Crec.JOTDAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_JOTDAT").Value, "-", ""), new string("0", Strings.Len(Crec.JOTDAT)));
//				Crec.JGMMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMDAY").Value, new string("0", Strings.Len(Crec.JGMMON)));
//				/// 종전근무월(GNMDAY로대체)
//				Crec.JEXMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JEXMON").Value, new string("0", Strings.Len(Crec.JEXMON)));
//				Crec.DUPMON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JMMMON").Value, new string("0", Strings.Len(Crec.DUPMON)));
//				Crec.GNMYER = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMYER").Value, new string("0", Strings.Len(Crec.GNMYER)));

//				/// 법정외 주현

//				/// 동일년도 중도정산후 퇴사자일경우 주현입사일대신 정산시작일입력함.
//				if (Strings.Left(Crec.JOTDAT, 4) == Strings.Left(Crec.INPDAT, 4)) {
//					Crec.FR2DAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_STRRET").Value, "-", ""), new string("0", Strings.Len(Crec.FR2DAT)));
//				} else {
//					Crec.FR2DAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_INPDAT").Value, "-", ""), new string("0", Strings.Len(Crec.FR2DAT)));
//				}
//				Crec.TO2DAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_ENDRET").Value, "-", ""), new string("0", Strings.Len(Crec.TO2DAT)));
//				///법정외근속월수가 없을경우 법정 근속내용삽입
//				if (oRecordSet.Fields.Item("U_GN2MON").Value == 0 | (Strings.Left(Crec.JOTDAT, 4) == Strings.Left(Crec.INPDAT, 4))) {
//					Crec.GN2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMMON").Value, new string("0", Strings.Len(Crec.GN2MON)));
//					Crec.EX2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_EXPMON").Value, new string("0", Strings.Len(Crec.EX2MON)));
//					Crec.GN2YER = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMYER").Value, new string("0", Strings.Len(Crec.GNMYER)));
//				} else {
//					Crec.GN2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GN2MON").Value, new string("0", Strings.Len(Crec.GN2MON)));
//					Crec.EX2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_EX2MON").Value, new string("0", Strings.Len(Crec.EX2MON)));
//					Crec.GN2YER = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GN2YER").Value, new string("0", Strings.Len(Crec.GN2YER)));
//				}
//				/// 법정외 종전
//				Crec.JI2DAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("ST2RET").Value, "-", ""), new string("0", Strings.Len(Crec.JI2DAT)));
//				//전  입사일자
//				Crec.JO2DAT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Replace(oRecordSet.Fields.Item("U_JOTDAT").Value, "-", ""), new string("0", Strings.Len(Crec.JO2DAT)));
//				/// 종전 법정외근속월수가 없을경우 법정 근속내용삽입
//				if (oRecordSet.Fields.Item("U_JO2MON").Value == 0) {
//					Crec.JO2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GNMDAY").Value, new string("0", Strings.Len(Crec.JO2MON)));
//					//    근속월수
//					Crec.JE2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JEXMON").Value, new string("0", Strings.Len(Crec.JE2MON)));
//					//    제외월수(2007년)
//					Crec.JB2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JMMMON").Value, new string("0", Strings.Len(Crec.JB2MON)));
//				} else {
//					Crec.JO2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JO2MON").Value, new string("0", Strings.Len(Crec.JO2MON)));
//					//    근속월수
//					Crec.JE2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JE2MON").Value, new string("0", Strings.Len(Crec.JE2MON)));
//					//    제외월수(2007년)
//					Crec.JB2MON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JB2MON").Value, new string("0", Strings.Len(Crec.JB2MON)));
//				}

//				/// 법정
//				Crec.JS1RET = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_RETPAY").Value) - Conversion.Val(oRecordSet.Fields.Item("U_SH2SUR").Value), new string("0", Strings.Len(Crec.JS1RET)));
//				Crec.JS1GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_RETGON").Value) - Conversion.Val(oRecordSet.Fields.Item("U_SH2GON").Value), new string("0", Strings.Len(Crec.JS1GON)));
//				Crec.JS1STD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_TAXSTD").Value) - Conversion.Val(oRecordSet.Fields.Item("U_SH2GWA").Value), new string("0", Strings.Len(Crec.JS1STD)));
//				Crec.JS1YSD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_YTXSTD").Value) - Conversion.Val(oRecordSet.Fields.Item("U_SH2YAG").Value), new string("0", Strings.Len(Crec.JS1YSD)));
//				Crec.JS1YST = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_YSANTX").Value) - Conversion.Val(oRecordSet.Fields.Item("U_SH2YAS").Value), new string("0", Strings.Len(Crec.JS1YST)));
//				Crec.JS1SAN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_SANTAX").Value) - Conversion.Val(oRecordSet.Fields.Item("U_JS2SAN").Value), new string("0", Strings.Len(Crec.JS1SAN)));
//				Crec.JS1FRN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Conversion.Val(oRecordSet.Fields.Item("U_TAXGON").Value) - Conversion.Val(oRecordSet.Fields.Item("U_JS2TAX").Value), new string("0", Strings.Len(Crec.JS1FRN)));
//				///2008년추가 세액(외국납부)공제
//				/// 법정외
//				Crec.JS2RET = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2SUR").Value, new string("0", Strings.Len(Crec.JS2RET)));
//				Crec.JS2GON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2GON").Value, new string("0", Strings.Len(Crec.JS2GON)));
//				Crec.JS2STD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2GWA").Value, new string("0", Strings.Len(Crec.JS2STD)));
//				Crec.JS2YSD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2YAG").Value, new string("0", Strings.Len(Crec.JS2YSD)));
//				Crec.JS2YST = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SH2YAS").Value, new string("0", Strings.Len(Crec.JS2YST)));
//				Crec.JS2SAN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JS2SAN").Value, new string("0", Strings.Len(Crec.JS2SAN)));
//				Crec.JS2FRN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JS2TAX").Value, new string("0", Strings.Len(Crec.JS2FRN)));
//				///2008년추가 세액(외국납부)공제
//				/// 계
//				Crec.RETPAY = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_RETPAY").Value, new string("0", Strings.Len(Crec.RETPAY)));
//				Crec.RETGON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_RETGON").Value, new string("0", Strings.Len(Crec.RETGON)));
//				Crec.TAXSTD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_TAXSTD").Value, new string("0", Strings.Len(Crec.TAXSTD)));
//				Crec.YTXSTD = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_YTXSTD").Value, new string("0", Strings.Len(Crec.YTXSTD)));
//				Crec.YSANTX = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_YSANTX").Value, new string("0", Strings.Len(Crec.YSANTX)));
//				Crec.SANTAX = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_SANTAX").Value, new string("0", Strings.Len(Crec.SANTAX)));
//				Crec.TAXGON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_TAXGON").Value, new string("0", Strings.Len(Crec.TAXGON)));
//				///2008년추가 세액(외국납부)공제
//				/// 결정
//				Crec.GULGAB = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GULGAB").Value, new string("0", Strings.Len(Crec.GULGAB)));
//				Crec.GULJUM = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GULJUM").Value, new string("0", Strings.Len(Crec.GULJUM)));
//				Crec.GULNON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.GULNON)));
//				Crec.GULTOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_GULGAB").Value + oRecordSet.Fields.Item("U_GULJUM").Value, new string("0", Strings.Len(Crec.GULTOT)));
//				/// 종전
//				Crec.JONGAB = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JONGAB").Value, new string("0", Strings.Len(Crec.JONGAB)));
//				Crec.JONJUM = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JONJUM").Value, new string("0", Strings.Len(Crec.JONJUM)));
//				Crec.JONNON = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Crec.JONNON)));
//				Crec.JONTOT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JONGAB").Value + oRecordSet.Fields.Item("U_JONJUM").Value, new string("0", Strings.Len(Crec.JONTOT)));
//				Crec.FILLER = Strings.Space(Strings.Len(Crec.FILLER));

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref Crec.RECGBN) + MDC_SetMod.sStr(ref Crec.DTAGBN) + MDC_SetMod.sStr(ref Crec.TAXCOD) + MDC_SetMod.sStr(ref Crec.SQNNBR) + MDC_SetMod.sStr(ref Crec.BUSNBR) + MDC_SetMod.sStr(ref Crec.BUSNUM) + MDC_SetMod.sStr(ref Crec.sangho) + MDC_SetMod.sStr(ref Crec.JONCNT) + MDC_SetMod.sStr(ref Crec.DWEGBN) + MDC_SetMod.sStr(ref Crec.RGNCOD) + MDC_SetMod.sStr(ref Crec.STRINT) + MDC_SetMod.sStr(ref Crec.ENDINT) + MDC_SetMod.sStr(ref Crec.MSTNAM) + MDC_SetMod.sStr(ref Crec.INTGBN) + MDC_SetMod.sStr(ref Crec.PERNBR) + MDC_SetMod.sStr(ref Crec.RETRES) + MDC_SetMod.sStr(ref Crec.TJKPAY) + MDC_SetMod.sStr(ref Crec.SUDAMT) + MDC_SetMod.sStr(ref Crec.BHMAMT) + MDC_SetMod.sStr(ref Crec.BHMAM1) + MDC_SetMod.sStr(ref Crec.TOTAMT) + MDC_SetMod.sStr(ref Crec.TOTAM1) + MDC_SetMod.sStr(ref Crec.BTXPAY) + MDC_SetMod.sStr(ref Crec.MYNACC) + MDC_SetMod.sStr(ref Crec.MYNTOT) + MDC_SetMod.sStr(ref Crec.MYNWON) + MDC_SetMod.sStr(ref Crec.MYNBUL) + MDC_SetMod.sStr(ref Crec.MYNGON) + MDC_SetMod.sStr(ref Crec.MYNYIL) + MDC_SetMod.sStr(ref Crec.JYNACC) + MDC_SetMod.sStr(ref Crec.JYNTOT) + MDC_SetMod.sStr(ref Crec.JYNWON) + MDC_SetMod.sStr(ref Crec.JYNBUL) + MDC_SetMod.sStr(ref Crec.JYNGON) + MDC_SetMod.sStr(ref Crec.JYITOT) + MDC_SetMod.sStr(ref Crec.SH1JIG) + MDC_SetMod.sStr(ref Crec.SH3JIG) + MDC_SetMod.sStr(ref Crec.SH1IYO) + MDC_SetMod.sStr(ref Crec.SH3IYO) + MDC_SetMod.sStr(ref Crec.SH1GIS) + MDC_SetMod.sStr(ref Crec.SH3GIS) + MDC_SetMod.sStr(ref Crec.SH1TIL) + MDC_SetMod.sStr(ref Crec.SH1SUR) + MDC_SetMod.sStr(ref Crec.SH1GON) + MDC_SetMod.sStr(ref Crec.SH1GWA) + MDC_SetMod.sStr(ref Crec.SH1YAG) + MDC_SetMod.sStr(ref Crec.SH1YAS) + MDC_SetMod.sStr(ref Crec.SH2JIG) + MDC_SetMod.sStr(ref Crec.SH4JIG) + MDC_SetMod.sStr(ref Crec.SH2IYO) + MDC_SetMod.sStr(ref Crec.SH4IYO) + MDC_SetMod.sStr(ref Crec.SH2GIS) + MDC_SetMod.sStr(ref Crec.SH4GIS) + MDC_SetMod.sStr(ref Crec.SH2TIL) + MDC_SetMod.sStr(ref Crec.SH2SUR) + MDC_SetMod.sStr(ref Crec.SH2GON) + MDC_SetMod.sStr(ref Crec.SH2GWA) + MDC_SetMod.sStr(ref Crec.SH2YAG) + MDC_SetMod.sStr(ref Crec.SH2YAS) + MDC_SetMod.sStr(ref Crec.INPDAT) + MDC_SetMod.sStr(ref Crec.OUTDAT) + MDC_SetMod.sStr(ref Crec.GNMMON) + MDC_SetMod.sStr(ref Crec.EXPMON) + MDC_SetMod.sStr(ref Crec.JIPDAT) + MDC_SetMod.sStr(ref Crec.JOTDAT) + MDC_SetMod.sStr(ref Crec.JGMMON) + MDC_SetMod.sStr(ref Crec.JEXMON) + MDC_SetMod.sStr(ref Crec.DUPMON) + MDC_SetMod.sStr(ref Crec.GNMYER) + MDC_SetMod.sStr(ref Crec.FR2DAT) + MDC_SetMod.sStr(ref Crec.TO2DAT) + MDC_SetMod.sStr(ref Crec.GN2MON) + MDC_SetMod.sStr(ref Crec.EX2MON) + MDC_SetMod.sStr(ref Crec.JI2DAT) + MDC_SetMod.sStr(ref Crec.JO2DAT) + MDC_SetMod.sStr(ref Crec.JO2MON) + MDC_SetMod.sStr(ref Crec.JE2MON) + MDC_SetMod.sStr(ref Crec.JB2MON) + MDC_SetMod.sStr(ref Crec.GN2YER) + MDC_SetMod.sStr(ref Crec.JS1RET) + MDC_SetMod.sStr(ref Crec.JS1GON) + MDC_SetMod.sStr(ref Crec.JS1STD) + MDC_SetMod.sStr(ref Crec.JS1YSD) + MDC_SetMod.sStr(ref Crec.JS1YST) + MDC_SetMod.sStr(ref Crec.JS1SAN) + MDC_SetMod.sStr(ref Crec.JS1FRN) + MDC_SetMod.sStr(ref Crec.JS2RET) + MDC_SetMod.sStr(ref Crec.JS2GON) + MDC_SetMod.sStr(ref Crec.JS2STD) + MDC_SetMod.sStr(ref Crec.JS2YSD) + MDC_SetMod.sStr(ref Crec.JS2YST) + MDC_SetMod.sStr(ref Crec.JS2SAN) + MDC_SetMod.sStr(ref Crec.JS2FRN) + MDC_SetMod.sStr(ref Crec.RETPAY) + MDC_SetMod.sStr(ref Crec.RETGON) + MDC_SetMod.sStr(ref Crec.TAXSTD) + MDC_SetMod.sStr(ref Crec.YTXSTD) + MDC_SetMod.sStr(ref Crec.YSANTX) + MDC_SetMod.sStr(ref Crec.SANTAX) + MDC_SetMod.sStr(ref Crec.TAXGON) + MDC_SetMod.sStr(ref Crec.GULGAB) + MDC_SetMod.sStr(ref Crec.GULJUM) + MDC_SetMod.sStr(ref Crec.GULNON) + MDC_SetMod.sStr(ref Crec.GULTOT) + MDC_SetMod.sStr(ref Crec.JONGAB) + MDC_SetMod.sStr(ref Crec.JONJUM) + MDC_SetMod.sStr(ref Crec.JONNON) + MDC_SetMod.sStr(ref Crec.JONTOT) + MDC_SetMod.sStr(ref Crec.FILLER));

//				Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + " 생성 완료.", true);
//				/// 필수입력 체크
//				if (Crec.DWEGBN == "0") {
//					Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-거주자구분이 누락되었습니다. 확인하여 주십시오.", ref false, ref true);
//					CheckC = Convert.ToString(true);
//				}
//				if (Crec.DWEGBN == "2" & string.IsNullOrEmpty(Strings.Trim(Crec.RGNCOD))) {
//					Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-비거주자일경우 거주지국코드는 필수입니다. 확인하여 주십시오.", ref false, ref true);
//					CheckC = Convert.ToString(true);
//				}
//				if (Crec.INTGBN == "1" & Strings.Len(Strings.Trim(Crec.PERNBR)) != 13) {
//					Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-주민등록번호를 확인하여 주십시오.", ref false, ref true);
//					CheckC = Convert.ToString(true);
//				}
//				if (Crec.RETRES == "0" | string.IsNullOrEmpty(Strings.Trim(Crec.RETRES))) {
//					Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(Crec.MSTNAM) + "-퇴직사유항목을 입력하여 주십시오.", ref false, ref true);
//					CheckC = Convert.ToString(true);
//				}

//				/// D레코드: 종전근무처 레코드
//				if (Conversion.Val(Crec.JONCNT) > 0) {
//					Drec.RECGBN = "D";
//					//DREC.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
//					Drec.DTAGBN = "25";
//					Drec.TAXCOD = Arec.TAXCOD;
//					Drec.SQNNBR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(Drec.SQNNBR)));

//					Drec.BUSNBR = Brec.BUSNBR;
//					Drec.FILLD1 = Strings.Space(Strings.Len(Drec.FILLD1));
//					Drec.PERNBR = Crec.PERNBR;

//					Drec.JONNAM = Strings.Trim(oRecordSet.Fields.Item("U_J01NAM").Value);
//					Drec.JONNBR = Strings.Replace(oRecordSet.Fields.Item("U_J01NBR").Value, "-", "");
//					Drec.RETAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JRET01").Value, new string("0", Strings.Len(Drec.RETAMT)));
//					Drec.SUDAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JSUD01").Value, new string("0", Strings.Len(Drec.SUDAMT)));
//					Drec.BHMAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JBHM01").Value, new string("0", Strings.Len(Drec.BHMAMT)));
//					Drec.BHMAM1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(0, new string("0", Strings.Len(Drec.BHMAM1)));
//					/// 컬럼을 안만들어놔서 일단 0으로 처리
//					Drec.TOTAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JTOT01").Value, new string("0", Strings.Len(Drec.TOTAMT)));
//					Drec.TOTAM1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_JADD01").Value, new string("0", Strings.Len(Drec.TOTAM1)));
//					Drec.BTXP01 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_BTXP01").Value, new string("0", Strings.Len(Drec.BTXP01)));
//					Drec.JONCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Crec.JONCNT, new string("0", Strings.Len(Drec.JONCNT)));
//					Drec.FILLER = Strings.Space(Strings.Len(Drec.FILLER));

//					FileSystem.PrintLine(1, MDC_SetMod.sStr(ref Drec.RECGBN) + MDC_SetMod.sStr(ref Drec.DTAGBN) + MDC_SetMod.sStr(ref Drec.TAXCOD) + MDC_SetMod.sStr(ref Drec.SQNNBR) + MDC_SetMod.sStr(ref Drec.BUSNBR) + MDC_SetMod.sStr(ref Drec.FILLD1) + MDC_SetMod.sStr(ref Drec.PERNBR) + MDC_SetMod.sStr(ref Drec.JONNAM) + MDC_SetMod.sStr(ref Drec.JONNBR) + MDC_SetMod.sStr(ref Drec.RETAMT) + MDC_SetMod.sStr(ref Drec.SUDAMT) + MDC_SetMod.sStr(ref Drec.BHMAMT) + MDC_SetMod.sStr(ref Drec.BHMAM1) + MDC_SetMod.sStr(ref Drec.TOTAMT) + MDC_SetMod.sStr(ref Drec.TOTAM1) + MDC_SetMod.sStr(ref Drec.BTXP01) + MDC_SetMod.sStr(ref Drec.JONCNT) + MDC_SetMod.sStr(ref Drec.FILLER));


//					/// 필수입력 체크
//					if (string.IsNullOrEmpty(Strings.Trim(Drec.JONNAM))) {
//						Matrix_AddRow("D레코드:종전근무처 법인명(상호)가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//						CheckC = Convert.ToString(true);
//					}
//					if (string.IsNullOrEmpty(Strings.Trim(Drec.JONNBR))) {
//						Matrix_AddRow("D레코드:종전근무처 사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//						CheckC = Convert.ToString(true);
//					}
//				}

//				oRecordSet.MoveNext();
//			}


//			if (Convert.ToBoolean(CheckC) == false) {
//				functionReturnValue = true;
//			} else {
//				functionReturnValue = false;
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			Error_Message:
//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("주(현)근무처 레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("D레코드(종전근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				Matrix_AddRow("C레코드오류: " + Err().Description, false);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void Matrix_AddRow(string MatrixMsg, ref bool Insert_YN = false, ref bool MatrixErr = false)
//		{
//			if (MatrixErr == true) {
//				oForm.DataSources.UserDataSources.Item("Col0").Value = "??";
//			} else {
//				oForm.DataSources.UserDataSources.Item("Col0").Value = "";
//			}
//			oForm.DataSources.UserDataSources.Item("Col1").Value = MatrixMsg;
//			if (Insert_YN == true) {
//				oMat1.AddRow();
//				MaxRow = MaxRow + 1;
//			}
//			oMat1.SetLineData(MaxRow);
//		}
//	}
//}
