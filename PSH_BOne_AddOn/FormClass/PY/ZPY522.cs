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
//	[System.Runtime.InteropServices.ProgId("ZPY522_NET.ZPY522")]
//	public class ZPY522
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : ZPY522.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 의료비 기부금 전산매체수록
//////  FormType       : 2000060522
//////  Create Date    : 2006.01.27
//////  Modified Date  :
//////  Creator        : Ham Mi Kyoung
//////  Modifier       :
//////  Copyright  (c) Morning Data
//////****************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbobsCOM.Recordset sRecordset;

//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;

//		private string oJsnYear;
//		private string DPTSTR;
//		private string DPTEND;
//		private string MSTCOD;
//		private string oFilePath;

//			//파  일  명
//		private Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString FILNAM = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(30);
//		private int MaxRow;
//		private short NEWCNT;
//		private string C_MSTCOD;
//		private string C_CLTCOD;

//		private string CLTCOD;
//			/// B레코드일련번호
//		private short BUSCNT;
//			/// B레코드총갯수
//		private short BUSTOT;

///// 기부금지급명세서 /
//		private struct H_A_Record
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
//			public char[] SANGHO;
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
//				//한글코드종류
//			public char[] HANCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//공      란
//			public char[] FILLER;
//		}
//		H_A_Record HArec;

//		private struct H_B_Record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] RECGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료 구분
//			public char[] DTAGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세무서
//			public char[] TAXCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련번호
//			public char[] BUSCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] BUSNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//상      호
//			public char[] SANGHO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//C레코드수
//			public char[] CRECNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//D레코드수
//			public char[] DRECNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부금액 총계
//			public char[] GBUAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//공제대상 금액 총계
//			public char[] GBUAM1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출대상기간
//			public char[] RNGCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(77), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 77)]
//				//공      란
//			public char[] FILLER;
//		}
//		H_B_Record HBrec;

//		private struct H_C_Record
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
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내외국인구분
//			public char[] INTGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//성      명
//			public char[] MSTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//유형코드
//			public char[] GBUCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//기부년도
//			public char[] GBUYER;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부금액
//			public char[] GBUAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//전년까지 공제된 금액
//			public char[] BEFAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//공제대상 금액
//			public char[] TARAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도 공제 금액
//			public char[] CURAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도 소멸 금액
//			public char[] DESAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//다음년도 이월 금액
//			public char[] NEXAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//기부금 조정명세 일련번호
//			public char[] GBUSEQ;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(25), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 25)]
//				//공란
//			public char[] FILLER;
//		}
//		H_C_Record HCrec;

//		private struct H_D_Record
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
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//유형코드
//			public char[] GBUCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부처 사업자번호
//			public char[] GBUNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//기부처 상호
//			public char[] GBUNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//관계코드
//			public char[] GWANGE;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내.외국인 구분
//			public char[] INTGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//기부자 성명
//			public char[] FAMNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부자 주민번호
//			public char[] FAMPER;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//연간 기부건수
//			public char[] GBUCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//연간 기부금액
//			public char[] GBUAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//해당년도 기부명세 일련번호
//			public char[] GBUSEQ;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(42), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 42)]
//				//공란
//			public char[] FILLER;
//		}
//		H_D_Record HDrec;

///// 의료비지급명세서 레코드 /
//		private struct CA_Record
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
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//제출  일자
//			public char[] PRTDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//자료제출자사업자번호
//			public char[] PRTBUS;
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
//			public char[] SANGHO;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민  번호
//			public char[] PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내외국인구분
//			public char[] INTGBN;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//성      명
//			public char[] MSTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//지급처사업자번호
//			public char[] JIGBUS;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//지급처상호
//			public char[] JIGNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//의료비증빙코드(2008.12추가)
//			public char[] MEDCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//지급(현금)건수
//			public char[] JIGCNT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//지급(현금)금액
//			public char[] JIGAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//지급 주민번호
//			public char[] JIGPER;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//지급 내외국인구분
//			public char[] JIGINT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//지급 해당여부
//			public char[] JIGCHK;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출대상기간(2009년추가)
//			public char[] RNGCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(19), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 19)]
//				//공란(2009년추가)
//			public char[] FILD01;
//		}
//		CA_Record CArec;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\ZPY522.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "ZPY522_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "ZPY522");
//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//			oForm.Freeze(true);
//			CreateItems();
//			oForm.Freeze(false);

//			oForm.EnableMenu(("1281"), false);
//			/// 찾기
//			oForm.EnableMenu(("1282"), true);
//			/// 추가
//			oForm.EnableMenu(("1284"), false);
//			/// 취소
//			oForm.EnableMenu(("1293"), false);
//			/// 행삭제
//			oForm.Update();
//			oForm.Visible = true;

//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			if ((oForm == null) == false) {
//				oForm.Freeze(false);
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//			}
//		}
////*******************************************************************
////// ItemEventHander
////*******************************************************************
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{

//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			switch (pval.EventType) {
//				//et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					if (pval.BeforeAction) {
//						/// ChooseBtn사원리스트
//						if (pval.ItemUID == "CBtn1") {
//							oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else if (pval.ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//							if (HeaderSpaceLineDel() == false) {
//								BubbleEvent = false;
//								return;
//							}
//							//UPGRADE_WARNING: oForm.Items(JsnGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (oForm.Items.Item("JsnGbn").Specific.Selected.Value == "1") {
//								if (File_MED_Create() == false) {
//									BubbleEvent = false;
//									return;
//								} else {
//									BubbleEvent = false;
//									oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//								}
//							} else {
//								if (File_GBU_Create() == false) {
//									BubbleEvent = false;
//									return;
//								} else {
//									BubbleEvent = false;
//									oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//								}

//							}
//						} else if (pval.ItemUID == "Btn1") {
//							oFilePath = My.MyProject.Forms.ZP_Form.vbGetBrowseDirectory(ref ZP_Form);
//							oForm.DataSources.UserDataSources.Item("Path").ValueEx = oFilePath;
//							BubbleEvent = false;
//							return;

//						}
//					} else {
//					}
//					break;
//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					if (pval.BeforeAction == false & pval.ItemChanged == true & (pval.ItemUID == "JsnYear" | pval.ItemUID == "MSTCOD")) {
//						FlushToItemValue(pval.ItemUID);
//					}
//					break;
//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					if (pval.BeforeAction == true & pval.ItemUID != "1000001" & pval.ItemUID != "2") {
//						///정산년도
//						if (Last_Item == "JsnYear") {
//							//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(Last_Item).Specific.Value))) {
//								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (MDC_SetMod.ChkYearMonth(ref Strings.Trim(Convert.ToString(oForm.Items.Item(Last_Item).Specific.Value)) + "01") == false) {
//									oForm.Items.Item(Last_Item).Update();
//									MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//									BubbleEvent = false;
//								}
//							}
//						} else if (Last_Item == "MSTCOD") {
//							//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(Last_Item).Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item(Last_Item).Specific.String) + "'", ref "") == true) {
//								oForm.Items.Item(Last_Item).Update();
//								MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					if (pval.BeforeAction == true & pval.ItemUID == "JsnYear" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Len(Strings.Trim(oForm.Items.Item(pval.ItemUID).Specific.String)) < 4) {
//							//UPGRADE_WARNING: oForm.Items(pval.ItemUid).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item(pval.ItemUID).Specific.Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item(pval.ItemUID).Specific.Value, "2000");
//						}
//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (MDC_SetMod.ChkYearMonth(ref Strings.Trim(Convert.ToString(oForm.Items.Item(pval.ItemUID).Specific.Value)) + "01") == false) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("정산년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					} else if (pval.BeforeAction == true & pval.ItemUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String)) & MDC_SetMod.Value_ChkYn(ref "[@PH_PY001A]", ref "Code", ref "'" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.String) + "'", ref "") == true) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("사원번호를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					}
//					break;
//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					if (Last_Item == "Mat1") {
//						if (pval.Row > 0) {
//							Last_Item = pval.ItemUID;
//						}
//					} else {
//						Last_Item = pval.ItemUID;
//					}
//					break;
//				//et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//					}
//					break;
//			}

//			return;
//			Raise_FormItemEvent_Error:
//			///////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
////*******************************************************************
////// MenuEventHander
////*******************************************************************
//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{

//			if (pval.BeforeAction == true) {
//				return;
//			}

//			switch (pval.MenuUID) {
//				case "1287":
//					/// 복제
//					break;
//				case "1281":
//				case "1282":
//					oForm.Items.Item("JsnYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					break;
//				case "1288": // TODO: to "1291"
//					break;
//				case "1293":
//					break;
//			}
//			return;
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			int i = 0;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Raise_FormDataEvent_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		private void CreateItems()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbouiCOM.ComboBox oCombo1 = null;
//			SAPbouiCOM.ComboBox oCombo2 = null;
//			SAPbouiCOM.OptionBtn oOption = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.Column oColumn = null;
//			string sQry = null;
//			SAPbouiCOM.Matrix Mat1 = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.DataSources.UserDataSources.Add("JsnYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			/// 생성년도
//			oForm.DataSources.UserDataSources.Add("JsnGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 생성구분
//			oForm.DataSources.UserDataSources.Add("JSNTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 기간코드
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 지점
//			oForm.DataSources.UserDataSources.Add("DptStr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			/// 부서코드
//			oForm.DataSources.UserDataSources.Add("DptEnd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
//			oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			oForm.DataSources.UserDataSources.Add("EmpID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oForm.DataSources.UserDataSources.Add("PRTDAT", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//			oEdit = oForm.Items.Item("JsnYear").Specific;
//			oEdit.DataBind.SetBound(true, "", "JsnYear");
//			oEdit = oForm.Items.Item("MSTCOD").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTCOD");
//			oEdit = oForm.Items.Item("MSTNAM").Specific;
//			oEdit.DataBind.SetBound(true, "", "MSTNAM");
//			oEdit = oForm.Items.Item("EmpID").Specific;
//			oEdit.DataBind.SetBound(true, "", "EmpID");
//			oEdit = oForm.Items.Item("Path").Specific;
//			oEdit.DataBind.SetBound(true, "", "Path");
//			oEdit = oForm.Items.Item("PRTDAT").Specific;
//			oEdit.DataBind.SetBound(true, "", "PRTDAT");

//			//// 생성구분
//			oCombo1 = oForm.Items.Item("JsnGbn").Specific;
//			oCombo1.ValidValues.Add("1", "의료비 지급명세서");
//			oCombo1.ValidValues.Add("2", "기부금 지급명세서");
//			oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			/// 전체
//			//// 생성구분
//			oCombo1 = oForm.Items.Item("JSNTYP").Specific;
//			oCombo1.ValidValues.Add("1", "연간(01.01~12.31)지급분");
//			oCombo1.ValidValues.Add("2", "폐업에 의한 수시 제출분");
//			oCombo1.ValidValues.Add("3", "수시 분할제출분");
//			oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			/// 전체

//			//// 자료 제출자
//			oCombo1 = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = " SELECT T0.U_WCHCLT, MAX(T1.NAME)  FROM [@PH_PY005A] T0 INNER JOIN [@PH_PY005A] T1 ON T0.U_WCHCLT = T1.CODE GROUP BY T0.U_WCHCLT  ORDER BY T0.U_WCHCLT"
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			while (!(oRecordSet.EoF)) {
//				oCombo1.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			if (oCombo1.ValidValues.Count > 0) {
//				oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				/// 전체
//			}
//			//// 부서
//			oCombo1 = oForm.Items.Item("DptStr").Specific;
//			oCombo2 = oForm.Items.Item("DptEnd").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			oCombo1.ValidValues.Add("-1", "모두");
//			oCombo2.ValidValues.Add("-1", "모두");
//			while (!(oRecordSet.EoF)) {
//				oCombo1.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oCombo2.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//				oRecordSet.MoveNext();
//			}
//			oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oCombo2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			oMat1 = oForm.Items.Item("Mat1").Specific;

//			oForm.DataSources.UserDataSources.Add("Col0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
//			oForm.DataSources.UserDataSources.Add("Col1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

//			oColumn = oMat1.Columns.Item("Col0");
//			oColumn.DataBind.SetBound(true, "", "Col0");

//			oColumn = oMat1.Columns.Item("Col1");
//			oColumn.DataBind.SetBound(true, "", "Col1");

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oOption 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOption = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: Mat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			Mat1 = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oOption 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOption = null;
//			//UPGRADE_NOTE: oCombo1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo1 = null;
//			//UPGRADE_NOTE: oCombo2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo2 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: Mat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			Mat1 = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}
//		private bool File_GBU_Create()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string oStr = null;
//			string sQry = null;

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
//			/// B레코드 일련번호
//			BUSTOT = 0;
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oJsnYear = oForm.Items.Item("JsnYear").Specific.Value;
//			/// 파일경로설정
//			if (string.IsNullOrEmpty(oFilePath))
//				oFilePath = "C:\\EOSDATA";
//			oFilePath = (Strings.Right(oFilePath, 1) == "\\" ? oFilePath : oFilePath + "\\");
//			oStr = MDC_SetMod.CreateFolder(ref Strings.Trim(oFilePath));
//			if (!string.IsNullOrEmpty(Strings.Trim(oStr))) {
//				ErrNum = 5;
//				goto Error_Message;
//			}

//			/// 기부금 제출자(대리인) 레코드
//			if (File_GBU_Create_ARecord() == false) {
//				ErrNum = 2;
//				goto Error_Message;
//			}

//			FileSystem.FileClose(1);
//			FileSystem.FileOpen(1, FILNAM.Value, OpenMode.Output);
//			/// A레코드: 기부금 원천징수의무자별 집계 레코드
//			FileSystem.PrintLine(1, MDC_SetMod.sStr(ref HArec.RECGBN) + MDC_SetMod.sStr(ref HArec.DTAGBN) + MDC_SetMod.sStr(ref HArec.TAXCOD) + MDC_SetMod.sStr(ref HArec.PRTDAT) + MDC_SetMod.sStr(ref HArec.RPTGBN) + MDC_SetMod.sStr(ref HArec.TAXAGE) + MDC_SetMod.sStr(ref HArec.HOMTID) + MDC_SetMod.sStr(ref HArec.PGMCOD) + MDC_SetMod.sStr(ref HArec.BUSNBR) + MDC_SetMod.sStr(ref HArec.SANGHO) + MDC_SetMod.sStr(ref HArec.DAMDPT) + MDC_SetMod.sStr(ref HArec.DAMNAM) + MDC_SetMod.sStr(ref HArec.DAMTEL) + MDC_SetMod.sStr(ref HArec.BUSCNT) + MDC_SetMod.sStr(ref HArec.HANCOD) + MDC_SetMod.sStr(ref HArec.FILLER));

//			Matrix_AddRow("제출자 레코드 생성 완료!", true);

//			/// B레코드: 기부금 집계 레코드 /***********************************************/
//			sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] ";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " WHERE  U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "' ORDER BY Code";
//			sRecordset.DoQuery(sQry);
//			while (!(sRecordset.EoF)) {
//				/// B레코드: 기부금 집계 레코드 /***********************************************/
//				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HBrec.TAXCOD = sRecordset.Fields.Item("U_TAXCODE").Value;
//				HBrec.BUSNBR = Strings.Replace(sRecordset.Fields.Item("U_BUSNUM").Value, "-", "");
//				HBrec.SANGHO = Strings.Trim(sRecordset.Fields.Item("U_CLTNAME").Value);

//				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CLTCOD = sRecordset.Fields.Item(0).Value;

//				switch (File_GBU_Create_Brecord()) {
//					case 0:
//						Matrix_AddRow(CLTCOD + "- 징수의무자의 집계 레코드 생성 완료!", true);
//						/// C레코드: 기부금 주(현)근무처 레코드 /***********************************************/
//						NEWCNT = 1;
//						if (File_GBU_Create_CRecord() == false) {
//							ErrNum = 4;
//							goto Error_Message;
//						}
//						Matrix_AddRow(CLTCOD + "- 징수의무자의 데이터 레코드" + NEWCNT + "건 생성 완료!", true);
//						break;
//					case 1:
//						ErrNum = 3;
//						goto Error_Message;
//						break;
//					case 2:
//						//// 해당자사에 기부금내역자료가 없으면 B,C,D레코드 생성을 건너뜀
//						break;

//				}
//				///
//				sRecordset.MoveNext();
//			}
//			FileSystem.FileClose(1);
//			oForm.DataSources.UserDataSources.Item("Path").Value = FILNAM.Value;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			functionReturnValue = true;
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			return functionReturnValue;
//			Error_Message:
//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("A레코드(기부금 제출자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B레코드(기부금 원천징수의무자별 집계 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드(기부금 주(현)근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("CreateFolder Error : " + oStr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("File_GBU_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private bool File_GBU_Create_ARecord()
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
//			//// 자사정보만 있고, 기부금자료가 없어서 C,D레코드 없이 B레코드만 생성되는 경우가 있으므로
//			//// 기부금 자료가 없으면 B레코드가 생성되지 않도록 함
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 " + "WHERE U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "' " + "AND Code IN (SELECT U_CLTCOD FROM [@ZPY505H] WHERE U_JSNYER = '" + oJsnYear + "')";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				BUSTOT = oRecordSet.Fields.Item(0).Value;
//			}
//			if (Conversion.Val(Convert.ToString(BUSTOT)) == 0) {
//				ErrNum = 3;
//				goto Error_Message;
//			}
//			/// 업체정보
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "'";

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
//				FILNAM.Value = oFilePath + "H" + Strings.Mid(BUSNUM, 1, 7) + "." + Strings.Mid(BUSNUM, 8, 3);
//				/// A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				HArec.RECGBN = "A";
//				//
//				HArec.DTAGBN = "27";
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HArec.TAXCOD = oRecordSet.Fields.Item("U_TAXCODE").Value;
//				HArec.PRTDAT = PRTDAT;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HArec.RPTGBN = oRecordSet.Fields.Item("U_TaxDGbn").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HArec.TAXAGE = oRecordSet.Fields.Item("U_TaxDCode").Value;
//				HArec.HOMTID = Strings.Trim(oRecordSet.Fields.Item("U_HOMETID").Value);
//				HArec.PGMCOD = "9000";
//				HArec.BUSNBR = Strings.Replace(oRecordSet.Fields.Item("U_TAXDBUS").Value, "-", "");
//				HArec.SANGHO = Strings.Trim(oRecordSet.Fields.Item("U_TAXDNAM").Value);
//				HArec.DAMDPT = Strings.Trim(oRecordSet.Fields.Item("U_CHGDPT").Value);
//				HArec.DAMNAM = Strings.Trim(oRecordSet.Fields.Item("U_CHGNAME").Value);
//				HArec.DAMTEL = Strings.Trim(oRecordSet.Fields.Item("U_CHGTEL").Value);
//				HArec.BUSCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(BUSTOT, new string("0", Strings.Len(HArec.BUSCNT)));
//				/// 원천징수의무자수
//				HArec.HANCOD = "101";
//				HArec.FILLER = Strings.Space(Strings.Len(HArec.FILLER));

//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.TAXCOD))) {
//					Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.RPTGBN))) {
//					Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.BUSNBR))) {
//					Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.SANGHO))) {
//					Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.DAMDPT))) {
//					Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.DAMNAM))) {
//					Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(HArec.DAMTEL))) {
//					Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (Convert.ToDouble(HArec.BUSCNT) == 0) {
//					Matrix_AddRow("A레코드:신고내역이 존재하는 B레코드가 없습니다. 확인하여 주십시오.", ref true, ref true);
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

//////------------------------------------------------------------
////// 반환값 리스트
////// 0 : 정상적으로 레코드 생성
////// 1 : 에러
////// 2 : B ~ E 레코드 생성 안함
//////------------------------------------------------------------
//		private short File_GBU_Create_Brecord()
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
//			sQry = "EXEC ZPY522 'B', '" + oJsnYear + "', '" + CLTCOD + "', '" + DPTSTR + "', '" + DPTEND + "','" + MSTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else if (oRecordSet.Fields.Item("CRECNT").Value == 0) {
//				ErrNum = 2;
//				goto Error_Message;
//			} else {
//				/// B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				BUSCNT = BUSCNT + 1;

//				HBrec.RECGBN = "B";
//				HBrec.DTAGBN = "27";
//				HBrec.BUSCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(BUSCNT, new string("0", Strings.Len(HBrec.BUSCNT)));
//				/// 원천징수의무자수 일련번호

//				HBrec.CRECNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("CRECNT").Value, new string("0", Strings.Len(HBrec.CRECNT)));
//				//C Record수
//				HBrec.DRECNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("DRECNT").Value, new string("0", Strings.Len(HBrec.DRECNT)));
//				//D Record수
//				HBrec.GBUAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUAMT").Value, new string("0", Strings.Len(HBrec.GBUAMT)));
//				//기부금액 총계
//				HBrec.GBUAM1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUAM1").Value, new string("0", Strings.Len(HBrec.GBUAM1)));
//				//공제대상 금액 총계
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HBrec.RNGCOD = oForm.Items.Item("JSNTYP").Specific.Selected.Value;
//				HBrec.FILLER = Strings.Space(Strings.Len(HBrec.FILLER));

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref HBrec.RECGBN) + MDC_SetMod.sStr(ref HBrec.DTAGBN) + MDC_SetMod.sStr(ref HBrec.TAXCOD) + MDC_SetMod.sStr(ref HBrec.BUSCNT) + MDC_SetMod.sStr(ref HBrec.BUSNBR) + MDC_SetMod.sStr(ref HBrec.SANGHO) + MDC_SetMod.sStr(ref HBrec.CRECNT) + MDC_SetMod.sStr(ref HBrec.DRECNT) + MDC_SetMod.sStr(ref HBrec.GBUAMT) + MDC_SetMod.sStr(ref HBrec.GBUAM1) + MDC_SetMod.sStr(ref HBrec.RNGCOD) + MDC_SetMod.sStr(ref HBrec.FILLER));

//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(HBrec.BUSNBR))) {
//					Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckB = Convert.ToString(true);
//				}
//			}

//			if (Convert.ToBoolean(CheckB) == false) {
//				functionReturnValue = 0;
//			} else {
//				functionReturnValue = 1;
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
//			} else if (ErrNum == 2) {
//				functionReturnValue = 2;
//			} else {
//				Matrix_AddRow("B레코드오류: " + Err().Description, false);
//				functionReturnValue = 1;
//			}
//			return functionReturnValue;

//		}
//		private bool File_GBU_Create_CRecord()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckC = null;
//			string BefMSTCOD = null;
//			int RecCNT = 0;

//			CheckC = Convert.ToString(false);
//			///체크필요유무

//			ErrNum = 0;
//			RecCNT = 0;
//			C_MSTCOD = "";
//			C_CLTCOD = "";

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 사원 정보 조회

//			BefMSTCOD = "";
//			sQry = "EXEC ZPY522 'C', '" + oJsnYear + "', '" + CLTCOD + "', '" + DPTSTR + "', '" + DPTEND + "','" + MSTCOD + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				//        ErrNum = 1
//				//        GoTo error_Message
//			}

//			while (!(oRecordSet.EoF)) {

//				/// D레코드: 해당년도 기부명세 레코드
//				if (!string.IsNullOrEmpty(BefMSTCOD) & BefMSTCOD != Strings.Trim(oRecordSet.Fields.Item("MSTCOD").Value)) {
//					if (File_GBU_Create_DRecord() == false) {
//						ErrNum = 2;
//						goto Error_Message;
//					}
//					NEWCNT = NEWCNT + 1;
//					RecCNT = 0;
//				}

//				C_MSTCOD = Strings.Trim(oRecordSet.Fields.Item("MSTCOD").Value);
//				C_CLTCOD = Strings.Trim(oRecordSet.Fields.Item("CLTCOD").Value);

//				/// 기부금 기부금 조정명세 레코드 /
//				/// C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//				HCrec.RECGBN = "C";
//				/// 레코드구분
//				HCrec.DTAGBN = "27";
//				/// 자료구분
//				HCrec.TAXCOD = HBrec.TAXCOD;
//				/// 세무서
//				HCrec.SQNNBR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(HCrec.SQNNBR)));
//				/// 일련번호

//				HCrec.BUSNBR = HBrec.BUSNBR;
//				/// 원천징수의무자 사업자번호
//				HCrec.PERNBR = Strings.Replace(oRecordSet.Fields.Item("PERNBR").Value, "-", "");
//				/// 소득자 주민번호
//				HCrec.INTGBN = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("INTGBN").Value, new string("0", Strings.Len(HCrec.INTGBN)));
//				/// 내.외국인 구분
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HCrec.MSTNAM = oRecordSet.Fields.Item("MSTNAM").Value;
//				/// 성명

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HCrec.GBUCOD = oRecordSet.Fields.Item("GBUCOD").Value;
//				/// 유형코드
//				HCrec.GBUYER = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUYER").Value, new string("0", Strings.Len(HCrec.GBUYER)));
//				/// 기부년도
//				HCrec.GBUAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUAMT").Value, new string("0", Strings.Len(HCrec.GBUAMT)));
//				/// 기부금액
//				HCrec.BEFAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("BEFAMT").Value, new string("0", Strings.Len(HCrec.BEFAMT)));
//				/// 전년까지 공제된 금액
//				HCrec.TARAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("TARAMT").Value, new string("0", Strings.Len(HCrec.TARAMT)));
//				/// 공제대상금액
//				HCrec.CURAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("CURAMT").Value, new string("0", Strings.Len(HCrec.CURAMT)));
//				/// 당해년도 공제금액
//				HCrec.DESAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("DESAMT").Value, new string("0", Strings.Len(HCrec.DESAMT)));
//				/// 당해년도 소멸금액
//				HCrec.NEXAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("NEXAMT").Value, new string("0", Strings.Len(HCrec.NEXAMT)));
//				/// 당해년도 이월금액
//				RecCNT = RecCNT + 1;
//				HCrec.GBUSEQ = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(RecCNT, new string("0", Strings.Len(HCrec.GBUSEQ)));
//				/// 소득자별 레코드 순번
//				HCrec.FILLER = Strings.Space(Strings.Len(HCrec.FILLER));

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref HCrec.RECGBN) + MDC_SetMod.sStr(ref HCrec.DTAGBN) + MDC_SetMod.sStr(ref HCrec.TAXCOD) + MDC_SetMod.sStr(ref HCrec.SQNNBR) + MDC_SetMod.sStr(ref HCrec.BUSNBR) + MDC_SetMod.sStr(ref HCrec.PERNBR) + MDC_SetMod.sStr(ref HCrec.INTGBN) + MDC_SetMod.sStr(ref HCrec.MSTNAM) + MDC_SetMod.sStr(ref HCrec.GBUCOD) + MDC_SetMod.sStr(ref HCrec.GBUYER) + MDC_SetMod.sStr(ref HCrec.GBUAMT) + MDC_SetMod.sStr(ref HCrec.BEFAMT) + MDC_SetMod.sStr(ref HCrec.TARAMT) + MDC_SetMod.sStr(ref HCrec.CURAMT) + MDC_SetMod.sStr(ref HCrec.DESAMT) + MDC_SetMod.sStr(ref HCrec.NEXAMT) + MDC_SetMod.sStr(ref HCrec.GBUSEQ) + MDC_SetMod.sStr(ref HCrec.FILLER));


//				Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(HCrec.MSTNAM) + " 생성 완료.", true);
//				/// 필수입력 체크
//				if (HCrec.INTGBN == "1" & Strings.Len(Strings.Trim(HCrec.PERNBR)) != 13) {
//					Matrix_AddRow("C레코드:" + C_MSTCOD + Strings.Trim(HCrec.MSTNAM) + "-주민등록번호를 확인하여 주십시오.", ref false, ref true);
//					CheckC = Convert.ToString(true);
//				}

//				BefMSTCOD = Strings.Trim(oRecordSet.Fields.Item("MSTCOD").Value);

//				oRecordSet.MoveNext();
//			}
//			/// D레코드: 해당년도 기부명세 레코드
//			if (File_GBU_Create_DRecord() == false) {
//				ErrNum = 2;
//				goto Error_Message;
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

//		private bool File_GBU_Create_DRecord()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckD = null;
//			short RecCNT = 0;

//			CheckD = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// 종전근무지정보
//			sQry = "EXEC ZPY522 'D', '" + oJsnYear + "', '" + C_CLTCOD + "', '" + DPTSTR + "', '" + DPTEND + "','" + C_MSTCOD + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				functionReturnValue = true;
//				return functionReturnValue;
//			}
//			RecCNT = 0;
//			while (!(oRecordSet.EoF)) {
//				/// D Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//				RecCNT = RecCNT + 1;
//				/// 사원별 해당년도 기부명세 일련번호
//				HDrec.RECGBN = "D";
//				///레코드구분
//				HDrec.DTAGBN = "27";
//				///자료구분
//				HDrec.TAXCOD = HArec.TAXCOD;
//				///세무서
//				HDrec.SQNNBR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(HDrec.SQNNBR)));
//				///소득자 일련번호

//				HDrec.BUSNBR = HBrec.BUSNBR;
//				///원천징수의무자 사업자등록번호
//				HDrec.PERNBR = HCrec.PERNBR;
//				///소득자 주민번호

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				HDrec.GBUCOD = oRecordSet.Fields.Item("GBUCOD").Value;
//				///유형코드
//				HDrec.GBUNBR = Strings.Trim(Strings.Replace(oRecordSet.Fields.Item("GBUNBR").Value, "-", ""));
//				///기부처 사업자번호
//				HDrec.GBUNAM = Strings.Trim(oRecordSet.Fields.Item("GBUNAM").Value);
//				///기부처명
//				HDrec.GWANGE = Strings.Trim(oRecordSet.Fields.Item("GWANGE").Value);
//				///기부자 관계코드
//				HDrec.INTGBN = Strings.Trim(oRecordSet.Fields.Item("INTGBN").Value);
//				///기부자 내.외국인
//				HDrec.FAMNAM = Strings.Trim(oRecordSet.Fields.Item("FAMNAM").Value);
//				///기부자명
//				HDrec.FAMPER = Strings.Trim(Strings.Replace(oRecordSet.Fields.Item("FAMPER").Value, "-", ""));
//				///기부자 주민번호
//				HDrec.GBUCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUCNT").Value, new string("0", Strings.Len(HDrec.GBUCNT)));
//				///기부 건수
//				HDrec.GBUAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("GBUAMT").Value, new string("0", Strings.Len(HDrec.GBUAMT)));
//				///기부 금액
//				HDrec.GBUSEQ = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(RecCNT, new string("0", Strings.Len(HDrec.GBUSEQ)));
//				///소득자별 레코드 순번
//				HDrec.FILLER = Strings.Space(Strings.Len(HDrec.FILLER));

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref HDrec.RECGBN) + MDC_SetMod.sStr(ref HDrec.DTAGBN) + MDC_SetMod.sStr(ref HDrec.TAXCOD) + MDC_SetMod.sStr(ref HDrec.SQNNBR) + MDC_SetMod.sStr(ref HDrec.BUSNBR) + MDC_SetMod.sStr(ref HDrec.PERNBR) + MDC_SetMod.sStr(ref HDrec.GBUCOD) + MDC_SetMod.sStr(ref HDrec.GBUNBR) + MDC_SetMod.sStr(ref HDrec.GBUNAM) + MDC_SetMod.sStr(ref HDrec.GWANGE) + MDC_SetMod.sStr(ref HDrec.INTGBN) + MDC_SetMod.sStr(ref HDrec.FAMNAM) + MDC_SetMod.sStr(ref HDrec.FAMPER) + MDC_SetMod.sStr(ref HDrec.GBUCNT) + MDC_SetMod.sStr(ref HDrec.GBUAMT) + MDC_SetMod.sStr(ref HDrec.GBUSEQ) + MDC_SetMod.sStr(ref HDrec.FILLER));

//				oRecordSet.MoveNext();
//			}

//			Matrix_AddRow("D레코드:" + C_MSTCOD + Strings.Trim(HCrec.MSTNAM) + " 당해년도 기부명세 생성 완료.", true);

//			if (Convert.ToBoolean(CheckD) == false) {
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

//			Matrix_AddRow("D레코드오류: " + Err().Description, false);
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private bool File_MED_Create()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string oStr = null;
//			string sQry = null;
//			string CLTNAM = null;

//			sRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			ErrNum = 0;
//			/// Question
//			if (MDC_Globals.Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			oMat1.Clear();
//			MaxRow = 0;
//			///
//			/// 파일경로설정
//			if (string.IsNullOrEmpty(oFilePath))
//				oFilePath = "C:\\EOSDATA";
//			oFilePath = (Strings.Right(oFilePath, 1) == "\\" ? oFilePath : oFilePath + "\\");
//			oStr = MDC_SetMod.CreateFolder(ref Strings.Trim(oFilePath));
//			if (!string.IsNullOrEmpty(Strings.Trim(oStr))) {
//				ErrNum = 5;
//				goto Error_Message;
//			}

//			/// 기부금 제출자(대리인) 레코드
//			if (File_MED_Create_ARecord() == false) {
//				ErrNum = 2;
//				goto Error_Message;
//			}
//			///
//			FileSystem.FileClose(1);
//			FileSystem.FileOpen(1, FILNAM.Value, OpenMode.Output);
//			sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] T0 ";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " WHERE U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "' ORDER BY CODE";
//			sRecordset.DoQuery(sQry);
//			NEWCNT = 0;
//			//// 일련번호를
//			while (!(sRecordset.EoF)) {
//				/// 원천징수의무자정보
//				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.TAXCOD = sRecordset.Fields.Item("U_TAXCODE").Value;
//				CArec.BUSNBR = Strings.Replace(sRecordset.Fields.Item("U_BUSNUM").Value, "-", "");
//				CArec.SANGHO = Strings.Trim(sRecordset.Fields.Item("U_CLTNAME").Value);
//				//UPGRADE_WARNING: sRecordset.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CLTCOD = sRecordset.Fields.Item(0).Value;
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CLTNAM = MDC_SetMod.Get_ReData("Name", "Code", "[@PH_PY005A]", "'" + Strings.Trim(CLTCOD) + "'");
//				/// 의료비명세 레코드
//				if (File_MED_Create_CARecord() == false) {
//					ErrNum = 3;
//					goto Error_Message;
//				}

//				Matrix_AddRow("신고의무자 : " + CLTCOD + "-" + CLTNAM + " 의 데이터 레코드" + NEWCNT + "건 생성 완료!", true);
//				///
//				sRecordset.MoveNext();
//			}
//			FileSystem.FileClose(1);
//			oForm.DataSources.UserDataSources.Item("Path").Value = FILNAM.Value;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			functionReturnValue = true;
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			return functionReturnValue;
//			Error_Message:
//			/////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: sRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			sRecordset = null;
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("제출자자료 조회를 실패하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("의료비명세 파일생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("CreateFolder Error : " + oStr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("File_MED_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private bool File_MED_Create_ARecord()
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
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 WHERE U_WCHCLT = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				BUSTOT = oRecordSet.Fields.Item(0).Value;
//			}
//			if (Conversion.Val(Convert.ToString(BUSTOT)) == 0) {
//				ErrNum = 3;
//				goto Error_Message;
//			}

//			/// 업체정보
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" + oForm.Items.Item("CLTCOD").Specific.Selected.Value + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				/// 파일명
//				BUSNUM = Strings.Replace(oRecordSet.Fields.Item("U_BUSNUM").Value, "-", "");
//				/// 징수의무자서업자번호
//				if (Strings.Len(Strings.Trim(BUSNUM)) != 10) {
//					ErrNum = 2;
//					goto Error_Message;
//				}
//				CArec.PRTBUS = Strings.Replace(oRecordSet.Fields.Item("U_TAXDBUS").Value, "-", "");
//				//제출자사업자번호
//				CArec.PRTDAT = PRTDAT;
//				//제출  일자
//				CArec.HOMTID = Strings.Trim(oRecordSet.Fields.Item("U_HOMETID").Value);
//				//홈텍스ID
//				CArec.PGMCOD = "9000";
//				//세무프로그램코드
//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.TAXCOD))) {
//					Matrix_AddRow("세무서코드가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.BUSNBR))) {
//					Matrix_AddRow("징수의무자의 사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.PRTBUS))) {
//					Matrix_AddRow("자료제출자의 사업자번호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.SANGHO))) {
//					Matrix_AddRow("상호가 누락되었습니다. 입력하여 주십시오.", ref true, ref true);
//					CheckA = Convert.ToString(true);
//				}
//				FILNAM.Value = oFilePath + "CA" + Strings.Mid(BUSNUM, 1, 7) + "." + Strings.Mid(BUSNUM, 8, 3);
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


//		private bool File_MED_Create_CARecord()
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

//			/// 사원별 의료비명세정보
//			sQry = " SELECT T0.*, ";
//			sQry = sQry + " T1.U_MSTCOD, ";
//			sQry = sQry + " T1.U_MSTNAM, ";
//			sQry = sQry + " ISNULL(T4.U_govID,'') AS MSTPER,";
//			sQry = sQry + " ISNULL(T4.U_INTGBN,1)   AS MSTINT";
//			sQry = sQry + " FROM [@ZPY506L] T0 ";
//			sQry = sQry + "      INNER JOIN [@ZPY506H] T1 ON T0.DocEntry = T1.DocEntry";
//			sQry = sQry + "      INNER JOIN [@PH_PY001A] T4 ON T1.U_MSTCOD = T4.Code";
//			sQry = sQry + "      INNER JOIN [@ZPY504H] T5 ON T1.U_MSTCOD = T5.U_MSTCOD AND T1.U_JSNYER = T5.U_JSNYER AND T1.U_CLTCOD = T5.U_CLTCOD";
//			sQry = sQry + " WHERE   T1.U_JSNYER = '" + oJsnYear + "'";
//			sQry = sQry + " AND     ISNULL(T1.U_CLTCOD, '') = '" + Strings.Trim(CLTCOD) + "'";
//			sQry = sQry + " AND     T4.U_TeamCode BETWEEN " + "'" + DPTSTR + "'" + " And " + "'" + DPTEND + "'";
//			sQry = sQry + " AND     T1.U_MSTCOD LIKE " + "N'" + Strings.Trim(MSTCOD) + "'";
//			sQry = sQry + " AND     ISNULL(T5.U_PILMED,0) >= 2000000";
//			sQry = sQry + " ORDER BY T0.DocEntry";
//			oRecordSet.DoQuery(sQry);
//			//    If oRecordSet.RecordCount = 0 Then  '// 신고자의 원천신고의무자가 1건이상일경우 기부금이 0건인 의무자자료 존재가능성있어서..주석처리함.
//			//        ErrNum = 1
//			//        GoTo error_Message
//			//    End If
//			while (!(oRecordSet.EoF)) {
//				NEWCNT = NEWCNT + 1;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_MSTCOD = oRecordSet.Fields.Item("U_MSTCOD").Value;
//				/// CA Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
//				CArec.RECGBN = "A";
//				CArec.DTAGBN = "26";
//				CArec.SQNNBR = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(CArec.SQNNBR)));
//				/// 일련번호

//				CArec.PERNBR = Strings.Replace(oRecordSet.Fields.Item("MSTPER").Value, "-", "");
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.INTGBN = oRecordSet.Fields.Item("MSTINT").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.MSTNAM = oRecordSet.Fields.Item("U_MSTNAM").Value;
//				CArec.JIGBUS = Strings.Replace(oRecordSet.Fields.Item("U_MEDNBR").Value, "-", "");
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.JIGNAM = oRecordSet.Fields.Item("U_MEDNAM").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.MEDCOD = oRecordSet.Fields.Item("U_MEDCOD").Value;
//				//        CArec.CADCNT = Format$(oRecordSet.Fields("U_MEDCNT2").Value, String$(Len(CArec.CADCNT), "0"))
//				//        CArec.CADAMT = Format$(oRecordSet.Fields("U_MEDAMT2").Value, String$(Len(CArec.CADAMT), "0"))
//				//        CArec.JIGCOD = oRecordSet.Fields("U_GWANGE").Value
//				CArec.JIGCNT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MEDCNT").Value, new string("0", Strings.Len(CArec.JIGCNT)));
//				CArec.JIGAMT = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("U_MEDAMT").Value, new string("0", Strings.Len(CArec.JIGAMT)));

//				CArec.JIGPER = Strings.Replace(oRecordSet.Fields.Item("U_PERNBR").Value, "-", "");
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.JIGINT = oRecordSet.Fields.Item("U_INTGBN").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.JIGCHK = oRecordSet.Fields.Item("U_DAECHK").Value;
//				CArec.FILD01 = Strings.Space(Strings.Len(CArec.FILD01));
//				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				CArec.RNGCOD = oForm.Items.Item("JSNTYP").Specific.Selected.Value;


//				FileSystem.PrintLine(1, CArec.RECGBN + CArec.DTAGBN + CArec.TAXCOD + CArec.SQNNBR + CArec.PRTDAT + CArec.PRTBUS + CArec.HOMTID + CArec.PGMCOD + CArec.BUSNBR + MDC_SetMod.sStr(ref CArec.SANGHO) + CArec.PERNBR + CArec.INTGBN + MDC_SetMod.sStr(ref CArec.MSTNAM) + CArec.JIGBUS + MDC_SetMod.sStr(ref CArec.JIGNAM) + CArec.MEDCOD + CArec.JIGCNT + CArec.JIGAMT + CArec.JIGPER + CArec.JIGINT + CArec.JIGCHK + CArec.RNGCOD + CArec.FILD01);


//				/// 필수입력 체크
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.JIGBUS)) & Strings.Trim(CArec.MEDCOD) != "1") {
//					Matrix_AddRow(C_MSTCOD + Strings.Trim(CArec.MSTNAM) + "지급처사업자등록번호가 누락되었습니다.", ref true, ref true);
//					CheckC = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.MEDCOD))) {
//					Matrix_AddRow(C_MSTCOD + Strings.Trim(CArec.MSTNAM) + "증빙코드가 누락되었습니다.", ref true, ref true);
//					CheckC = Convert.ToString(true);
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(CArec.JIGPER))) {
//					Matrix_AddRow(C_MSTCOD + Strings.Trim(CArec.MSTNAM) + "지급대상자주민등록번호가 누락되었습니다.", ref true, ref true);
//					CheckC = Convert.ToString(true);
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("의료비명세자료 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				Matrix_AddRow("CA레코드오류: " + CArec.BUSNBR + "사업장의 의료비명세자료 존재하지 않습니다.", false);
//			} else {
//				Matrix_AddRow("CA레코드오류: " + Err().Description, false);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void FlushToItemValue(string oUID, ref int oRow = 0)
//		{
//			ZPAY_g_EmpID MstInfo = default(ZPAY_g_EmpID);

//			switch (oUID) {
//				case "JsnYear":
//					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oJsnYear = oForm.Items.Item(oUID).Specific.String;
//					if (string.IsNullOrEmpty(Strings.Trim(oJsnYear))) {
//						MDC_Globals.ZPAY_GBL_JSNYER.Value = oJsnYear;
//					}
//					break;
//				case "MSTCOD":
//					//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oForm.Items.Item(oUID).Specific.String)) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.Items.Item(oUID).Specific.String = "";
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = "";
//						oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = "";
//					} else {
//						//UPGRADE_WARNING: oForm.Items(oUID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oForm.Items.Item(oUID).Specific.String = Strings.UCase(oForm.Items.Item(oUID).Specific.String);
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						//UPGRADE_WARNING: MstInfo 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MstInfo = MDC_SetMod.Get_EmpID_InFo(ref oForm.Items.Item(oUID).Specific.String);
//						oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = MstInfo.MSTNAM;
//						oForm.DataSources.UserDataSources.Item("EmpID").ValueEx = MstInfo.EmpID;
//					}
//					oForm.Items.Item("MSTNAM").Update();
//					oForm.Items.Item("EmpID").Update();
//					break;
//			}
//			oForm.Items.Item(oUID).Update();
//		}

//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//저장할 데이터의 유효성을 점검한다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;

//			ErrNum = 0;
//			/// 필수Check
//			//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			/// 정산년도
//			if (MDC_SetMod.ChkYearMonth(ref oForm.Items.Item("JsnYear").Specific.String + "01") == false) {
//				ErrNum = 1;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(JsnGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("JsnGbn").Specific.Selected == null) {
//				ErrNum = 2;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (oForm.Items.Item("CLTCOD").Specific.Selected == null) {
//				ErrNum = 3;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("PRTDAT").Specific.String))) {
//				ErrNum = 4;
//				goto HeaderSpaceLineDel;
//			}

//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTSTR = oForm.Items.Item("DptStr").Specific.Selected.Value;
//			if (DPTSTR == "-1")
//				DPTSTR = "00000001";
//			//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DPTEND = oForm.Items.Item("DptEnd").Specific.Selected.Value;
//			if (DPTEND == "-1")
//				DPTEND = "ZZZZZZZZ";
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
//				MSTCOD = "%";

//			functionReturnValue = true;
//			return functionReturnValue;
//			HeaderSpaceLineDel:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("기간코드는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("지점은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
