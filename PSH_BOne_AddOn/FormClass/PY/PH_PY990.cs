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
//	internal class PH_PY990
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : PH_PY990.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 기부금지급명세서자료 전산매체수록
//////  FormType       :
//////  Create Date    : 2014.02.03
//////  Modified Date  : 2016.01.10
//////  Creator        : NGY
//////  Modifier       :
//////  Copyright  (c) Poongsan Holdings
//////****************************************************************************


//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;
//		private SAPbobsCOM.Recordset sRecordset;
//		private SAPbouiCOM.Matrix oMat1;
//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string Last_Item;

//		private string CLTCOD;
//		private string yyyy;
//		private string HtaxID;
//		private string TeamName;
//		private string Dname;
//		private string Dtel;
//		private string DocDate;
//		private string oFilePath;

//			//파  일  명
//		private Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString FILNAM = new Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(30);
//		private int MaxRow;
//			/// B레코드일련번호
//		private short BUSCNT;
//			/// B레코드총갯수
//		private short BUSTOT;

//		private short NEWCNT;
//		private short OLDCNT;
//		private string C_SAUP;
//		private string C_YYYY;
//		private string C_SABUN;
//		private string E_BUYCNT;
//		private string C_BUYCNT;


////2013년기준 180 BYTE
////2016년기준 190 BYTE
//		private struct A_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] A001;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료구분
//			public char[] A002;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세무서
//			public char[] A003;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//제출일자
//			public char[] A004;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출자구분 (1;세무대리인, 2;법인, 3;개인)
//			public char[] A005;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//세무대리인
//			public char[] A006;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//홈텍스ID
//			public char[] A007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//세무프로그램코드
//			public char[] A008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] A009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//법인명(상호)
//			public char[] A010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//담당자부서
//			public char[] A011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//담당자성명
//			public char[] A012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(15), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 15)]
//				//담당자전화번호
//			public char[] A013;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//신고의무자수
//			public char[] A014;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//한글코드종류
//			public char[] A015;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(12), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 12)]
//				//공란
//			public char[] A016;
//		}
//		A_record A_rec;


//		private struct B_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] B001;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료구분
//			public char[] B002;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세무서
//			public char[] B003;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련번호
//			public char[] B004;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] B005;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//법인명(상호)
//			public char[] B006;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//C레코드수
//			public char[] B007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//D레코드수
//			public char[] B008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부금액합계
//			public char[] B009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//공제대상금액합계
//			public char[] B010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출대상기간코드
//			public char[] B011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(87), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 87)]
//				//공란
//			public char[] B012;
//		}
//		B_record B_rec;


//		private struct C_record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] C001;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료구분
//			public char[] C002;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세무서
//			public char[] C003;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련번호
//			public char[] C004;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] C005;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//소득자주민등록번호
//			public char[] C006;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내,외국인코드
//			public char[] C007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//성명
//			public char[] C008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//유형코드
//			public char[] C009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//기부년도
//			public char[] C010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부금액
//			public char[] C011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//전년까지공제된금액
//			public char[] C012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//공제대상금액
//			public char[] C013;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도공제금액 필요경비 '0'  2016
//			public char[] C014;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도공제금액세액(소득)공제금액  2016
//			public char[] C015;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도에공제받지못한금액_소멸금액
//			public char[] C016;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//해당년도에공제받지못한금액_이월금액
//			public char[] C017;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//기부조정명세일련번호
//			public char[] C018;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(22), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//				//공란
//			public char[] C019;
//		}
//		C_record C_rec;

//		private struct D_Record
//		{
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//레코드구분
//			public char[] D001;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//자료구분
//			public char[] D002;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//				//세무서
//			public char[] D003;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//				//일련번호
//			public char[] D004;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자등록번호
//			public char[] D005;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민등록번호
//			public char[] D006;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//유형코드
//			public char[] D007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부처-사업자등록번호
//			public char[] D008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//기부처-법인명(상호)
//			public char[] D009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//관계
//			public char[] D010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내,외국인코드
//			public char[] D011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//성명
//			public char[] D012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민등록번호
//			public char[] D013;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//건수
//			public char[] D014;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//금액
//			public char[] D015;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//공제대상기부금액  2016
//			public char[] D016;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//기부장려금신청금액  2016
//			public char[] D017;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//				//해당연도기부명세일련번호
//			public char[] D018;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(26), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 26)]
//				//공란
//			public char[] D019;
//		}
//		D_Record D_rec;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY990.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "PH_PY990_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY990");
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
//						//                If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn사원리스트
//						//                    oForm.Items("MSTCOD").CLICK ct_Regular
//						//                    Sbo_Application.ActivateMenuItem ("7425")
//						//                    BubbleEvent = False
//						//                Else

//						if (pval.ItemUID == "Btn01") {
//							if (HeaderSpaceLineDel() == false) {
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

//						}
//					} else {
//					}
//					break;

//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							switch (pval.ItemUID) {
//								////사업장이 바뀌면
//								case "CLTCOD":
//									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.Value) + "'";
//									oRecordSet.DoQuery(sQry);
//									//UPGRADE_WARNING: oForm.Items(HtaxID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("HtaxID").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_HomeTId").Value);
//									//UPGRADE_WARNING: oForm.Items(TeamName).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("TeamName").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_ChgDpt").Value);
//									//UPGRADE_WARNING: oForm.Items(Dname).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("Dname").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_ChgName").Value);
//									//UPGRADE_WARNING: oForm.Items(Dtel).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("Dtel").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_ChgTel").Value);
//									break;

//							}
//						}
//					}
//					break;

//				//et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					break;

//				//et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					break;

//				//et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					break;

//				//et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
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

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			oCombo.DataBind.SetBound(true, "", "CLTCOD");
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;
//			//// 접속자에 따른 권한별 사업장 콤보박스세팅
//			MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			//UPGRADE_WARNING: oForm.Items(YYYY).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YYYY").Specific.String = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;
//			//년도 기본년도에서 - 1

//			oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			//제출일자
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//		}

//		private bool File_Create()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string oStr = null;
//			string sQry = null;

//			sRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//화면변수를 전역변수로 MOVE
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = Strings.Trim(oForm.Items.Item("YYYY").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			HtaxID = Strings.Trim(oForm.Items.Item("HtaxID").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TeamName = Strings.Trim(oForm.Items.Item("TeamName").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Dname = Strings.Trim(oForm.Items.Item("Dname").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Dtel = Strings.Trim(oForm.Items.Item("Dtel").Specific.Value);
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);

//			ErrNum = 0;

//			/// Question
//			if (MDC_Globals.Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			/// A RECORD 처리
//			if (File_Create_A_record() == false) {
//				ErrNum = 2;
//				goto Error_Message;
//			}

//			/// B RECORD 처리
//			if (File_Create_B_record() == false) {
//				ErrNum = 3;
//				goto Error_Message;
//			}

//			/// C RECORD 처리  D 처리
//			if (File_Create_C_record() == false) {
//				ErrNum = 4;
//				goto Error_Message;
//			}

//			FileSystem.FileClose(1);

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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("A레코드(제출자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B레코드(원천징수의무자별 집계 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드(기부금 조정명세 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//		private bool File_Create_A_record()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string PRTDAT = null;
//			string saup = null;
//			string CheckA = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			CheckA = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			/// A_RECORE QUERY
//			sQry = "EXEC PH_PY990_A '" + CLTCOD + "', '" + HtaxID + "', '" + TeamName + "', '" + Dname + "', '" + Dtel + "', '" + DocDate + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				// PATH및 파일이름 만들기
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				saup = oRecordSet.Fields.Item("A009").Value;
//				//사업자번호
//				oFilePath = "C:\\BANK\\H" + Strings.Mid(saup, 1, 7) + "." + Strings.Mid(saup, 8, 3);


//				//A RECORD MOVE

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A001 = oRecordSet.Fields.Item("A001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A002 = oRecordSet.Fields.Item("A002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A003 = oRecordSet.Fields.Item("A003").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A004 = oRecordSet.Fields.Item("A004").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A005 = oRecordSet.Fields.Item("A005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A006 = oRecordSet.Fields.Item("A006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A007 = oRecordSet.Fields.Item("A007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A008 = oRecordSet.Fields.Item("A008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A009 = oRecordSet.Fields.Item("A009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A010 = oRecordSet.Fields.Item("A010").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A011 = oRecordSet.Fields.Item("A011").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A012 = oRecordSet.Fields.Item("A012").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A013 = oRecordSet.Fields.Item("A013").Value;

//				A_rec.A014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("A014").Value, new string("0", Strings.Len(A_rec.A014)));
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A015 = oRecordSet.Fields.Item("A015").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A016 = oRecordSet.Fields.Item("A016").Value;

//				FileSystem.FileClose(1);
//				FileSystem.FileOpen(1, oFilePath, OpenMode.Output);
//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref A_rec.A001) + MDC_SetMod.sStr(ref A_rec.A002) + MDC_SetMod.sStr(ref A_rec.A003) + MDC_SetMod.sStr(ref A_rec.A004) + MDC_SetMod.sStr(ref A_rec.A005) + MDC_SetMod.sStr(ref A_rec.A006) + MDC_SetMod.sStr(ref A_rec.A007) + MDC_SetMod.sStr(ref A_rec.A008) + MDC_SetMod.sStr(ref A_rec.A009) + MDC_SetMod.sStr(ref A_rec.A010) + MDC_SetMod.sStr(ref A_rec.A011) + MDC_SetMod.sStr(ref A_rec.A012) + MDC_SetMod.sStr(ref A_rec.A013) + MDC_SetMod.sStr(ref A_rec.A014) + MDC_SetMod.sStr(ref A_rec.A015) + MDC_SetMod.sStr(ref A_rec.A016));

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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("귀속년도의 자사정보(A RECORD)가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				Matrix_AddRow("A레코드오류: " + Err().Description, ref false, ref true);
//			}

//			functionReturnValue = false;
//			return functionReturnValue;

//		}

//		private short File_Create_B_record()
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

//			/// B_RECORE QUERY
//			sQry = "EXEC PH_PY990_B '" + CLTCOD + "', '" + yyyy + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				//B RECORD MOVE

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B001 = oRecordSet.Fields.Item("B001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B002 = oRecordSet.Fields.Item("B002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B003 = oRecordSet.Fields.Item("B003").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B004 = oRecordSet.Fields.Item("B004").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B005 = oRecordSet.Fields.Item("B005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B006 = oRecordSet.Fields.Item("B006").Value;
//				B_rec.B007 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B007").Value, new string("0", Strings.Len(B_rec.B007)));
//				B_rec.B008 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B008").Value, new string("0", Strings.Len(B_rec.B008)));
//				B_rec.B009 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B009").Value, new string("0", Strings.Len(B_rec.B009)));
//				B_rec.B010 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B010").Value, new string("0", Strings.Len(B_rec.B010)));
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B011 = oRecordSet.Fields.Item("B011").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B012 = oRecordSet.Fields.Item("B012").Value;

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref B_rec.B001) + MDC_SetMod.sStr(ref B_rec.B002) + MDC_SetMod.sStr(ref B_rec.B003) + MDC_SetMod.sStr(ref B_rec.B004) + MDC_SetMod.sStr(ref B_rec.B005) + MDC_SetMod.sStr(ref B_rec.B006) + MDC_SetMod.sStr(ref B_rec.B007) + MDC_SetMod.sStr(ref B_rec.B008) + MDC_SetMod.sStr(ref B_rec.B009) + MDC_SetMod.sStr(ref B_rec.B010) + MDC_SetMod.sStr(ref B_rec.B011) + MDC_SetMod.sStr(ref B_rec.B012));

//			}

//			if (Convert.ToBoolean(CheckB) == false) {
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				functionReturnValue = 1;
//			} else {
//				Matrix_AddRow("B레코드오류: " + Err().Description, false);
//				functionReturnValue = 1;
//			}
//			return functionReturnValue;

//		}

//		private bool File_Create_C_record()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckC = null;
//			string PSABUN = null;
//			double OLDBIG = 0;
//			double PILTOT = 0;
//			short SCount = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			CheckC = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			/// C_RECORE QUERY
//			sQry = "EXEC PH_PY990_C '" + CLTCOD + "', '" + yyyy + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, false);

//			NEWCNT = 1;
//			SCount = 0;
//			//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PSABUN = oRecordSet.Fields.Item("sabun").Value;

//			while (!(oRecordSet.EoF)) {

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_SAUP = oRecordSet.Fields.Item("saup").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_YYYY = oRecordSet.Fields.Item("yyyy").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_SABUN = oRecordSet.Fields.Item("sabun").Value;

//				//C RECORD MOVE

//				SCount = SCount + 1;

//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C001 = oRecordSet.Fields.Item("C001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C002 = oRecordSet.Fields.Item("C002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C003 = oRecordSet.Fields.Item("C003").Value;
//				C_rec.C004 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(C_rec.C004)));
//				/// 일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C005 = oRecordSet.Fields.Item("C005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C006 = oRecordSet.Fields.Item("C006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C007 = oRecordSet.Fields.Item("C007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C008 = oRecordSet.Fields.Item("C008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C009 = oRecordSet.Fields.Item("C009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C010 = oRecordSet.Fields.Item("C010").Value;

//				C_rec.C011 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C011").Value, new string("0", Strings.Len(C_rec.C011)));
//				C_rec.C012 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C012").Value, new string("0", Strings.Len(C_rec.C012)));
//				C_rec.C013 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C013").Value, new string("0", Strings.Len(C_rec.C013)));
//				C_rec.C014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C014").Value, new string("0", Strings.Len(C_rec.C014)));
//				C_rec.C015 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C015").Value, new string("0", Strings.Len(C_rec.C015)));
//				C_rec.C016 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C016").Value, new string("0", Strings.Len(C_rec.C016)));
//				C_rec.C017 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C017").Value, new string("0", Strings.Len(C_rec.C017)));
//				C_rec.C018 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(SCount, new string("0", Strings.Len(C_rec.C018)));
//				/// 일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C019 = oRecordSet.Fields.Item("C019").Value;


//				//예제
//				//C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").Value, "-", "")

//				//OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").Value) + Val(oRecordSet.Fields("U_BIGWA3").Value) + Val(oRecordSet.Fields("U_BIGWA5").Value) _
//				//'        + Val(oRecordSet.Fields("U_BIGWA6").Value) + Val(oRecordSet.Fields("U_BIGWU3").Value)

//				//C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
//				//C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
//				//C_rec.FILLER = Space$(Len(C_rec.FILLER))
//				//C_rec.C022 = Format$(oRecordSet.Fields("C022").Value, , String$(Len(C_rec.C022), "0"))


//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref C_rec.C001) + MDC_SetMod.sStr(ref C_rec.C002) + MDC_SetMod.sStr(ref C_rec.C003) + MDC_SetMod.sStr(ref C_rec.C004) + MDC_SetMod.sStr(ref C_rec.C005) + MDC_SetMod.sStr(ref C_rec.C006) + MDC_SetMod.sStr(ref C_rec.C007) + MDC_SetMod.sStr(ref C_rec.C008) + MDC_SetMod.sStr(ref C_rec.C009) + MDC_SetMod.sStr(ref C_rec.C010) + MDC_SetMod.sStr(ref C_rec.C011) + MDC_SetMod.sStr(ref C_rec.C012) + MDC_SetMod.sStr(ref C_rec.C013) + MDC_SetMod.sStr(ref C_rec.C014) + MDC_SetMod.sStr(ref C_rec.C015) + MDC_SetMod.sStr(ref C_rec.C016) + MDC_SetMod.sStr(ref C_rec.C017) + MDC_SetMod.sStr(ref C_rec.C018) + MDC_SetMod.sStr(ref C_rec.C019));


//				oRecordSet.MoveNext();


//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";


//				if (oRecordSet.EoF) {
//					/// D레코드
//					if (File_Create_D_record() == false) {
//						ErrNum = 2;
//						goto Error_Message;
//					}

//				} else if (PSABUN != oRecordSet.Fields.Item("sabun").Value) {
//					/// D레코드
//					if (File_Create_D_record() == false) {
//						ErrNum = 2;
//						goto Error_Message;
//					}

//					NEWCNT = NEWCNT + 1;
//					/// 일련번호
//					SCount = 0;
//					//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					PSABUN = oRecordSet.Fields.Item("sabun").Value;

//				}

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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("D레코드 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				Matrix_AddRow("C레코드오류: " + Err().Description, false);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private bool File_Create_D_record()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			string CheckD = null;
//			short DCount = 0;

//			CheckD = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// D_RECORE QUERY
//			sQry = "EXEC PH_PY990_D '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";

//			oRecordSet.DoQuery(sQry);

//			DCount = 0;
//			while (!(oRecordSet.EoF)) {

//				//D RECORD MOVE
//				DCount = DCount + 1;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D001 = oRecordSet.Fields.Item("D001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D002 = oRecordSet.Fields.Item("D002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D003 = oRecordSet.Fields.Item("D003").Value;
//				D_rec.D004 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(C_rec.C004, new string("0", Strings.Len(D_rec.D004)));
//				/// C레코드의 일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D005 = oRecordSet.Fields.Item("D005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D006 = oRecordSet.Fields.Item("D006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D007 = oRecordSet.Fields.Item("D007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D008 = oRecordSet.Fields.Item("D008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D009 = oRecordSet.Fields.Item("D009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D010 = oRecordSet.Fields.Item("D010").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D011 = oRecordSet.Fields.Item("D011").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D012 = oRecordSet.Fields.Item("D012").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D013 = oRecordSet.Fields.Item("D013").Value;
//				D_rec.D014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D014").Value, new string("0", Strings.Len(D_rec.D014)));
//				D_rec.D015 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D015").Value, new string("0", Strings.Len(D_rec.D015)));
//				D_rec.D016 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D016").Value, new string("0", Strings.Len(D_rec.D016)));
//				D_rec.D017 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D017").Value, new string("0", Strings.Len(D_rec.D017)));
//				D_rec.D018 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DCount, new string("0", Strings.Len(D_rec.D018)));
//				///일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D019 = oRecordSet.Fields.Item("D019").Value;

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref D_rec.D001) + MDC_SetMod.sStr(ref D_rec.D002) + MDC_SetMod.sStr(ref D_rec.D003) + MDC_SetMod.sStr(ref D_rec.D004) + MDC_SetMod.sStr(ref D_rec.D005) + MDC_SetMod.sStr(ref D_rec.D006) + MDC_SetMod.sStr(ref D_rec.D007) + MDC_SetMod.sStr(ref D_rec.D008) + MDC_SetMod.sStr(ref D_rec.D009) + MDC_SetMod.sStr(ref D_rec.D010) + MDC_SetMod.sStr(ref D_rec.D011) + MDC_SetMod.sStr(ref D_rec.D012) + MDC_SetMod.sStr(ref D_rec.D013) + MDC_SetMod.sStr(ref D_rec.D014) + MDC_SetMod.sStr(ref D_rec.D015) + MDC_SetMod.sStr(ref D_rec.D016) + MDC_SetMod.sStr(ref D_rec.D017) + MDC_SetMod.sStr(ref D_rec.D018) + MDC_SetMod.sStr(ref D_rec.D019));

//				oRecordSet.MoveNext();
//			}

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

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("D레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				Matrix_AddRow("D레코드오류: " + Err().Description, false);
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
////화면변수 CHECK
//		private bool HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;

//			ErrNum = 0;
//			/// 필수Check
//			//UPGRADE_WARNING: oForm.Items(HtaxID).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(oForm.Items.Item("HtaxID").Specific.Value)) {
//				ErrNum = 1;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(TeamName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("TeamName").Specific.Value)) {
//				ErrNum = 2;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(Dname).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("Dname").Specific.Value)) {
//				ErrNum = 3;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(Dtel).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("Dtel").Specific.Value)) {
//				ErrNum = 4;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
//				ErrNum = 5;
//				goto HeaderSpaceLineDel;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			HeaderSpaceLineDel:
//			///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("홈텍스ID(5자리이상)를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("담당자부서는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 3) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("담당자성명은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 4) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("담당자전화번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else if (ErrNum == 5) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			}

//			functionReturnValue = false;
//			return functionReturnValue;
//		}
//	}
//}
