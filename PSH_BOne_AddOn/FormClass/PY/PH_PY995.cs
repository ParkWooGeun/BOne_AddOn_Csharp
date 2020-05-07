using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{

    /// <summary>
    /// 퇴직소득지급명세서자료 전산매체수록
    /// </summary>
    internal class PH_PY995 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        /// <summary>
        /// Form 호출
        /// </summary>
        public override void LoadForm()
        {
            string strXml = string.Empty;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY995.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY995_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY995");

                strXml = oXmlDoc.xml.ToString();
                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                CreateItems();

                oForm.EnableMenu("1281", false); //찾기
                oForm.EnableMenu("1282", true); //추가
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
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
                oForm.ActiveItem = "DocDate"; //제출일자로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                // 정산년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.Value = DateTime.Now.AddYears(-1).ToString("yyyy"); // 기본년도에서 - 1

                // 홈택스ID
                oForm.DataSources.UserDataSources.Add("HtaxID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("HtaxID").Specific.DataBind.SetBound(true, "", "HtaxID");

                // 담당자부서
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당자성명
                oForm.DataSources.UserDataSources.Add("Dname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("Dname").Specific.DataBind.SetBound(true, "", "Dname");

                // 담당자전화번호
                oForm.DataSources.UserDataSources.Add("Dtel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("Dtel").Specific.DataBind.SetBound(true, "", "Dtel");

                // 제출일자
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 신고파일 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create()
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string stringSpace = string.Empty;
            string CLTCOD, yyyy, HtaxID, TeamName, Dname, Dtel, DocDate = string.Empty;

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
                HtaxID = oForm.Items.Item("HtaxID").Specific.Value.ToString().Trim();
                TeamName = oForm.Items.Item("TeamName").Specific.Value.ToString().Trim();
                Dname = oForm.Items.Item("Dname").Specific.Value.ToString().Trim();
                Dtel = oForm.Items.Item("Dtel").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                errNum = 0;

                // Question
                if (PSH_Globals.SBO_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2)
                {
                    errNum = 1;
                    throw new Exception();
                }

                // A RECORD 처리
                if (File_Create_A_record(CLTCOD, HtaxID, TeamName, Dname, Dtel, DocDate) == false)
                {
                    errNum = 2;
                    throw new Exception();
                }
                // B RECORD 처리
                if (File_Create_B_record(CLTCOD, yyyy) == false)
                {
                    errNum = 3;
                    throw new Exception();
                }
                // C RECORD 처리
                if (File_Create_C_record(CLTCOD, yyyy) == false)
                {
                    errNum = 4;
                    throw new Exception();
                }
                FileSystem.FileClose(1);

                PSH_Globals.SBO_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A레코드(근로 제출자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B레코드(근로 원천징수의무자별 집계 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("C레코드(근로 주(현)근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    stringSpace = new string(' ', 10);
                    PSH_Globals.SBO_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + stringSpace + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }

            return functionReturnValue;
        }

        /// <summary>
        /// A레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_A_record(string pCLTCOD, string pHtaxID, string pTeamName, string pDname, string pDtel, string pDocDate)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry = string.Empty;
            string saup = string.Empty;
            string oFilePath = string.Empty; //파일 경로

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // A 제출자 레코드
            // 2013년기준 1400 BYTE
            // 2014년기준 1090 BYTE
            // 2016년기준 1110 BYTE 
            // 2017년기준 1110 BYTE

            string A001;  //  1     '레코드구분
            string A002;  //  2     '자료구분
            string A003;  //  3     '세무서
            string A004;  //  8     '제출일자
            string A005;  //  1     '제출자구분 (1;세무대리인, 2;법인, 3;개인)
            string A006;  //  6     '세무대리인관리번호
            string A007;  //  20    '홈텍스ID
            string A008;  //  4     '세무프로그램코드
            string A009;  //  10    '사업자번호
            string A010;  //  40    '법인명(상호)
            string A011;  //  30    '담당자부서
            string A012;  //  30    '담당자성명
            string A013;  //  15    '담당자전화번호
            string A014;  //  5     '신고의무자수
            string A015;  //  3     '한글코드종류
            string A016;  //  932   '공란

            try
            {
                //A_RECORE QUERY
                sQry = "EXEC PH_PY995_A '" + pCLTCOD + "', '" + pHtaxID + "', '" + pTeamName + "', '" + pDname + "', '" + pDtel + "', '" + pDocDate + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    //PATH및 파일이름 만들기
                    saup = oRecordSet.Fields.Item("A009").Value; //사업자번호
                    oFilePath = "C:\\BANK\\EA" + codeHelpClass.Mid(saup, 0, 7) + "." + codeHelpClass.Mid(saup, 7, 3);
                    FileSystem.FileClose(1);
                    FileSystem.FileOpen(1, oFilePath, OpenMode.Output);

                    // A RECORD MOVE
                    A001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A001").Value.ToString().Trim(), 1);
                    A002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A002").Value.ToString().Trim(), 2);
                    A003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A003").Value.ToString().Trim(), 3);
                    A004 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A004").Value.ToString().Trim(), 8);
                    A005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A005").Value.ToString().Trim(), 1);
                    A006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A006").Value.ToString().Trim(), 6);
                    A007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A007").Value.ToString().Trim(), 20);
                    A008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A008").Value.ToString().Trim(), 4);
                    A009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A009").Value.ToString().Trim(), 10);
                    A010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A010").Value.ToString().Trim(), 40);
                    A011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A011").Value.ToString().Trim(), 30);
                    A012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A012").Value.ToString().Trim(), 30);
                    A013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A013").Value.ToString().Trim(), 15);
                    A014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A014").Value.ToString().Trim(), 5, '0');
                    A015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A015").Value.ToString().Trim(), 3);
                    A016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A016").Value.ToString().Trim(), 932);

                    FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015 + A016);

                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속년도의 자사정보(A RECORD)가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// B레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_B_record(string pCLTCOD, string pyyyy)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // B 원천징수의무자별 집계 레코드
            string B001;    // 1     '레코드구분
            string B002;    // 2     '자료구분
            string B003;    // 3     '세무서
            string B004;    // 6     '일련번호
            string B005;    // 10    '사업자번호
            string B006;    // 40    '법인명(상호)
            string B007;    // 30    '대표자(성명)
            string B008;    // 13    '주민(법인)등록번호
            string B009;    // 1     '제출대상기간코드
            string B010;    // 7     '퇴직소득자(C레코드)수
            string B011;    // 7     '공란
            string B012;    // 14    '정산-과세대상퇴직금여합계
            string B013_1;  // 1     '부호
            string B013_2;  // 13    '신고대상소득세합계
            string B014;    // 13    '이연퇴직소득세액합계
            string B015_1;  // 1     '부호
            string B015_2;  // 13    '차감원천징수-소득세액합계
            string B016_1;  // 1     '부호
            string B016_2;  // 13    '차감원천징수-지방소득세액합계
            string B017_1;  // 1     '부호
            string B017_2;  // 13    '차감원천징수-농어촌특별세액합계
            string B018_1;  // 1     '부호
            string B018_2;  // 13    '차감원천징수세액-계합계
            string B019;    // 893   '공란

            try
            {
                // B_RECORE QUERY
                sQry = "EXEC PH_PY995_B '" + pCLTCOD + "', '" + pyyyy + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    // B RECORD MOVE
                    B001   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B001").Value.ToString().Trim(), 1);
                    B002   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B002").Value.ToString().Trim(), 2);
                    B003   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B003").Value.ToString().Trim(), 3);
                    B004   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B004").Value.ToString().Trim(), 6);
                    B005   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B005").Value.ToString().Trim(), 10);
                    B006   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B006").Value.ToString().Trim(), 40);
                    B007   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B007").Value.ToString().Trim(), 30);
                    B008   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B008").Value.ToString().Trim(), 13);
                    B009   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B009").Value.ToString().Trim(), 1);
                    B010   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B010").Value.ToString().Trim(), 7, '0');
                    B011   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B011").Value.ToString().Trim(), 7);
                    B012   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B012").Value.ToString().Trim(), 14, '0');
                    B013_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013_1").Value.ToString().Trim(), 1);
                    B013_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013_2").Value.ToString().Trim(), 13, '0');
                    B014   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B014").Value.ToString().Trim(), 13, '0');
                    B015_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B015_1").Value.ToString().Trim(), 1);
                    B015_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B015_2").Value.ToString().Trim(), 13, '0');
                    B016_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B016_1").Value.ToString().Trim(), 1);
                    B016_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B016_2").Value.ToString().Trim(), 13, '0');
                    B017_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B017_1").Value.ToString().Trim(), 1);
                    B017_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B017_2").Value.ToString().Trim(), 13, '0');
                    B018_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018_1").Value.ToString().Trim(), 1);
                    B018_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018_2").Value.ToString().Trim(), 13, '0');
                    B019   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B019").Value.ToString().Trim(), 893);

                    FileSystem.PrintLine(1, B001 + B002 + B003 + B004 + B005 + B006 + B007 + B008 + B009 + B010 + B011 + B012 + B013_1 + B013_2 + B014
                                          + B015_1 + B015_2 + B016_1 + B016_2 + B017_1 + B017_2 + B018_1 + B018_2 + B019);
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// C레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_C_record(string pCLTCOD, string pyyyy)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry = string.Empty;
            string C_SAUP, C_YYYY, C_SABUN = string.Empty;
            int NEWCNT = 0; //일련번호

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // C 레코드(퇴직소득자 레코드)
            string C001;   // 1     '레코드구분
            string C002;   // 2     '자료구분
            string C003;   // 3     '세무서
            string C004;   // 6     '일련번호
            string C005;   // 10    '사업자번호
            string C006;   // 1     '징수의무자구분 1.사업장 3.공적연금사업자
            // 소득자
            string C007;   // 1     '거주구분 1.거주자 2.비거주자
            string C008;   // 1     '내.외국인구분 1.내국인 9.외국인
            string C009;   // 2     '거주지국코드
            string C010;   // 30    '성명
            string C011;   // 13    '주민등록번호
            string C012;   // 1     '임원여부  1.여 2.부
            string C013;   // 8     '확정급여형퇴직연금제도가입일
            string C014;   // 11    '2011.12.31 퇴직금
            string C015;   // 8     '귀속연도시작연월일
            string C016;   // 8     '귀속연도종료연월일
            string C017;   // 1     '퇴직사유
            // 퇴직급여현황-중간지급등
            string C018;   // 40    '근무처명
            string C019;   // 10    '근무처사업자등록번호
            string C020;   // 11    '퇴직급여
            string C021;   // 11    '비과세퇴직급여
            string C022;   // 11    '과세대상퇴직급여
            // 퇴직급여현황-최종분
            string C023;   // 40    '근무처명
            string C024;   // 10    '근무처사업자등록번호
            string C025;   // 11    '퇴직급여
            string C026;   // 11    '비과세퇴직급여
            string C027;   // 11    '과세대상퇴직급여
            // 퇴직급여현황-정산
            string C028;   // 11    '퇴직급여
            string C029;   // 11    '비과세퇴직급여
            string C030;   // 11    '과세대상퇴직급여
            // 근속연수-중간지급등
            string C031;   // 8     '입사일
            string C032;   // 8     '기산일
            string C033;   // 8     '퇴사일
            string C034;   // 8     '지급일
            string C035;   // 4     '근속월수
            string C036;   // 4     '제외월수
            string C037;   // 4     '가산월수
            string C038;   // 4     '중복월수
            string C039;   // 4     '근속연수
            // 근속연수-최종
            string C040;   // 8     '입사일
            string C041;   // 8     '기산일
            string C042;   // 8     '퇴사일
            string C043;   // 8     '지급일
            string C044;   // 4     '근속월수
            string C045;   // 4     '제외월수
            string C046;   // 4     '가산월수
            string C047;   // 4     '중복월수
            string C048;   // 4     '근속연수
            // 근속연수-정산
            string C049;   // 8     '입사일
            string C050;   // 8     '기산일
            string C051;   // 8     '퇴사일
            string C052;   // 8     '지급일
            string C053;   // 4     '근속월수
            string C054;   // 4     '제외월수
            string C055;   // 4     '가산월수
            string C056;   // 4     '중복월수
            string C057;   // 4     '근속연수
            // 근속연수-안분-2012.12.31이전
            string C058;   // 8     '입사일
            string C059;   // 8     '기산일
            string C060;   // 8     '퇴사일
            string C061;   // 8     '지급일
            string C062;   // 4     '근속월수
            string C063;   // 4     '제외월수
            string C064;   // 4     '가산월수
            string C065;   // 4     '중복월수
            string C066;   // 4     '근속연수
            // 근속연수-안분-2013.01.01이후
            string C067;   // 8     '입사일
            string C068;   // 8     '기산일
            string C069;   // 8     '퇴사일
            string C070;   // 8     '지급일
            string C071;   // 4     '근속월수
            string C072;   // 4     '제외월수
            string C073;   // 4     '가산월수
            string C074;   // 4     '중복월수
            string C075;   // 4     '근속연수
            // 개정규정에따른계산방법-과세표준계산
            string C076;   // 11    '퇴직소득
            string C077;   // 11    '근속연수공제
            string C078;   // 11    '환산급여
            string C079;   // 11    '환산급여별공제
            string C080;   // 11    '퇴직소득과세표준
            // 개정규정에따른계산방법-세액계산
            string C081;   // 11    '환산산출세액
            string C082;   // 11    '산출세액
            // 종전규정에 따른 계산 방법 - 과세표준계산
            string C083;   // 11    '퇴직소득
            string C084;   // 11    '퇴직소득정률공제
            string C085;   // 11    '근속연수공제
            string C086;   // 11    '퇴직소득과세표준
            // 종전규정에 따른 계산 방법 - 세액계산 - 2012.12.31 이전
            string C087;   // 11    '과세표준안분
            string C088;   // 11    '연평균과세표준
            string C089;   // 11    '환산과세표준
            string C090;   // 11    '환산산출세액
            string C091;   // 11    '연평균산출세액
            string C092;   // 11    '산출세액
            // 종전규정에 따른 계산 방법 - 세액계산 - 2013.1.1 이후
            string C093;   // 11    '과세표준안분
            string C094;   // 11    '연평균과세표준
            string C095;   // 11    '환산과세표준
            string C096;   // 11    '환산산출세액
            string C097;   // 11    '연평균산출세액
            string C098;   // 11    '산출세액
            // 종전규정에 따른 계산 방법 - 세액계산 - 합계
            string C099;   // 11    '과세표준안분
            string C100;   // 11    '연평균과세표준
            string C101;   // 11    '환산과세표준
            string C102;   // 11    '환산산출세액
            string C103;   // 11    '연평균산출세액
            string C104;   // 11    '산출세액
            // 퇴직소득세액계산
            string C105;   // 4     '퇴직일이속하는과세연도
            string C106;   // 11    '퇴직소득세산출세액
            string C107;   // 11    '기납부(또는기과세이연)세액
            string C108_1; // 1     '부호
            string C108_2; // 11    '신고대상세액
            // 이연퇴직소득세액계산
            string C109_1; // 1     '부호
            string C109_2; // 11    '신고대상세액
            string C110;   // 11    '계좌입금금액_합계
            string C111;   // 11    '퇴직급여
            string C112;   // 11    '이연퇴직소득세
            // 납부명세-신고대상세액
            string C113_1; // 1     '부호
            string C113_2; // 11    '소득세
            string C114_1; // 1     '부호
            string C114_2; // 11    '지방소득세
            string C115_1; // 1     '부호
            string C115_2; // 11    '농어총특별세
            string C116_1; // 1     '부호
            string C116_2; // 11    '계
            // 납부명세-이연퇴직소득세
            string C117;   // 11    '소득세
            string C118;   // 11    '지방소득세
            string C119;   // 11    '농어촌특별세
            string C120;   // 11    '계
            // 납부명세-차감원천징수세액
            string C121_1; // 1     '부호
            string C121_2; // 11    '소득세
            string C122_1; // 1     '부호
            string C122_2; // 11    '지방소득세
            string C123_1; // 1     '부호
            string C123_2; // 11    '농어총특별세
            string C124_1; // 1     '부호
            string C124_2; // 11    '계
            string C125;   // 2     '공란

            try
            {
                // C_RECORE QUERY
                sQry = "EXEC PH_PY995_C '" + pCLTCOD + "', '" + pyyyy + "'";
                oRecordSet.DoQuery(sQry);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, false);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    NEWCNT = 0;
                    while (!oRecordSet.EoF)
                    {
                        NEWCNT = NEWCNT + 1; //일련번호
                        C_SAUP = oRecordSet.Fields.Item("saup").Value.ToString().Trim();
                        C_YYYY = oRecordSet.Fields.Item("yyyy").Value.ToString().Trim();
                        C_SABUN = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();

                        //C RECORD MOVE
                        C001   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C001").Value.ToString().Trim(), 1);
                        C002   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C002").Value.ToString().Trim(), 2);
                        C003   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C003").Value.ToString().Trim(), 3);
                        C004   = codeHelpClass.GetFixedLengthStringByte(NEWCNT.ToString(), 6, '0');
                        C005   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C005").Value.ToString().Trim(), 10);
                        C006   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C006").Value.ToString().Trim(), 1, '0');
                             
                        C007   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C007").Value.ToString().Trim(), 1);
                        C008   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C008").Value.ToString().Trim(), 1);
                        C009   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C009").Value.ToString().Trim(), 2);
                        C010   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C010").Value.ToString().Trim(), 30);
                        C011   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C011").Value.ToString().Trim(), 13);
                        C012   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C012").Value.ToString().Trim(), 1);
                        C013   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C013").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        // TABLE이 소수 6자리(numeric(19,6))라서 사사오입 시킴(C#은 기본으로 .5가 반올림 안됨 그래서 MidpointRounding.AwayFromZero 문을 씀)
                        C014 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C014").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C015   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C015").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        C016   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C016").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        C017   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C017").Value.ToString().Trim(), 1);
                                                                                                                                
                        C018   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C018").Value.ToString().Trim(), 40);
                        C019   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C019").Value.ToString().Trim(), 10);
                        C020   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C020").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C021   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C021").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C022   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C022").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                                
                        C023   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C023").Value.ToString().Trim(), 40);
                        C024   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C024").Value.ToString().Trim(), 10);
                        C025   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C025").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C026   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C026").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C027   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C027").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C028   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C028").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C029   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C029").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C030   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C030").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                            
                        C031   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C031").Value.ToString().Trim(), 8, '0');
                        C032   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C032").Value.ToString().Trim(), 8, '0');
                        C033   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C033").Value.ToString().Trim(), 8, '0');
                        C034   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C034").Value.ToString().Trim(), 8, '0');
                        C035   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C035").Value.ToString().Trim(), 4, '0');
                        C036   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C036").Value.ToString().Trim(), 4, '0');
                        C037   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C037").Value.ToString().Trim(), 4, '0');
                        C038   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C038").Value.ToString().Trim(), 4, '0');
                        C039   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C039").Value.ToString().Trim(), 4, '0');
                                                                                                            
                        C040   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C040").Value.ToString().Trim(), 8, '0');
                        C041   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C041").Value.ToString().Trim(), 8, '0');
                        C042   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C042").Value.ToString().Trim(), 8, '0');
                        C043   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C043").Value.ToString().Trim(), 8, '0');
                        C044   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C044").Value.ToString().Trim(), 4, '0');
                        C045   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C045").Value.ToString().Trim(), 4, '0');
                        C046   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C046").Value.ToString().Trim(), 4, '0');
                        C047   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C047").Value.ToString().Trim(), 4, '0');
                        C048   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C048").Value.ToString().Trim(), 4, '0');
                                                                                                           
                        C049   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C049").Value.ToString().Trim(), 8, '0');
                        C050   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C050").Value.ToString().Trim(), 8, '0');
                        C051   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C051").Value.ToString().Trim(), 8, '0');
                        C052   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C052").Value.ToString().Trim(), 8, '0');
                        C053   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C053").Value.ToString().Trim(), 4, '0');
                        C054   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C054").Value.ToString().Trim(), 4, '0');
                        C055   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C055").Value.ToString().Trim(), 4, '0');
                        C056   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C056").Value.ToString().Trim(), 4, '0');
                        C057   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C057").Value.ToString().Trim(), 4, '0');
                                                                                                                      
                        C058   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C058").Value.ToString().Trim(), 8, '0');
                        C059   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C059").Value.ToString().Trim(), 8, '0');
                        C060   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C060").Value.ToString().Trim(), 8, '0');
                        C061   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C061").Value.ToString().Trim(), 8, '0');
                        C062   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C062").Value.ToString().Trim(), 4, '0');
                        C063   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C063").Value.ToString().Trim(), 4, '0');
                        C064   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C064").Value.ToString().Trim(), 4, '0');
                        C065   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C065").Value.ToString().Trim(), 4, '0');
                        C066   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C066").Value.ToString().Trim(), 4, '0');
                                                                                                                      
                        C067   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C067").Value.ToString().Trim(), 8, '0');
                        C068   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C068").Value.ToString().Trim(), 8, '0');
                        C069   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C069").Value.ToString().Trim(), 8, '0');
                        C070   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C070").Value.ToString().Trim(), 8, '0');
                        C071   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C071").Value.ToString().Trim(), 4, '0');
                        C072   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C072").Value.ToString().Trim(), 4, '0');
                        C073   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C073").Value.ToString().Trim(), 4, '0');
                        C074   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C074").Value.ToString().Trim(), 4, '0');
                        C075   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C075").Value.ToString().Trim(), 4, '0');
                                                                                                                            
                        C076   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C076").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C077   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C077").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C078   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C078").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C079   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C079").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C080   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C080").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C081   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C081").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C082   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C082").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C083   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C083").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C084   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C084").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C085   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C085").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C086   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C086").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C087   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C087").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C088   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C088").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C089   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C089").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C090   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C090").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C091   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C091").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C092   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C092").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C093   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C093").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C094   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C094").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C095   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C095").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C096   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C096").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C097   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C097").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C098   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C098").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                   
                        C099   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C099").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C100   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C100").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C101   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C101").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C102   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C102").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C103   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C103").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C104   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C104").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                            
                        C105   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C105").Value.ToString().Trim(), 4, '0'); 
                        C106   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C106").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C107   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C107").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C108_1 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C108_1").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 1, '0');
                        C108_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C108_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                                  
                        C109_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C109_1").Value.ToString().Trim(), 1, '0');
                        C109_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C109_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C110   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C110").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C111   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C111").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C112   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C112").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                                
                        C113_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C113_1").Value.ToString().Trim(), 1, '0');
                        C113_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C113_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C114_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C114_1").Value.ToString().Trim(), 1, '0');
                        C114_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C114_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C115_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C115_1").Value.ToString().Trim(), 1, '0');
                        C115_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C115_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C116_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C116_1").Value.ToString().Trim(), 1, '0');
                        C116_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C116_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                                
                        C117   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C117").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C118   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C118").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C119   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C119").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0'); 
                        C120   = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C120").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                                                                                                                                
                        C121_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C121_1").Value.ToString().Trim(), 1, '0');
                        C121_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C121_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C122_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C122_1").Value.ToString().Trim(), 1, '0');
                        C122_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C122_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C123_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C123_1").Value.ToString().Trim(), 1, '0');
                        C123_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C123_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C124_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C124_1").Value.ToString().Trim(), 1, '0');
                        C124_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C124_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C125   = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C125").Value.ToString().Trim(), 2);

                        FileSystem.PrintLine(1, C001 + C002 + C003 + C004 + C005 + C006 + C007 + C008 + C009 + C010 + C011 + C012 + C013 + C014 + C015 + C016 + C017 + C018 + C019 + C020
                                              + C021 + C022 + C023 + C024 + C025 + C026 + C027 + C028 + C029 + C030 + C031 + C032 + C033 + C034 + C035 + C036 + C037 + C038 + C039 + C040
                                              + C041 + C042 + C043 + C044 + C045 + C046 + C047 + C048 + C049 + C050 + C051 + C052 + C053 + C054 + C055 + C056 + C057 + C058 + C059 + C060
                                              + C061 + C062 + C063 + C064 + C065 + C066 + C067 + C068 + C069 + C070 + C071 + C072 + C073 + C074 + C075 + C076 + C077 + C078 + C079 + C080
                                              + C081 + C082 + C083 + C084 + C085 + C086 + C087 + C088 + C089 + C090 + C091 + C092 + C093 + C094 + C095 + C096 + C097 + C098 + C099 + C100
                                              + C101 + C102 + C103 + C104 + C105 + C106 + C107 + C108_1 + C108_2 + C109_1 + C109_2 + C110 + C111 + C112 + C113_1 + C113_2 + C114_1 + C114_2 
                                              + C115_1 + C115_2 + C116_1 + C116_2 + C117 + C118 + C119 + C120 + C121_1 + C121_2 + C122_1 + C122_2 + C123_1 + C123_2 + C124_1 + C124_2 + C125);

                        // D 레코드 : 연금계좌입금명세 레코드 수록
                        if (Conversion.Val(C110) > 0)    // 계좌입금금액_합계  2016부터있음
                        {
                            if (File_Create_D_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                            {
                                errNum = 2;
                                throw new Exception();
                            }
                        }

                        oRecordSet.MoveNext();

                        ProgressBar01.Value = ProgressBar01.Value + 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";

                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                ProgressBar01.Stop();
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("C레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D레코드 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// D 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_D_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // D 연금계좌입금명세 레코드 수록
            string D001;  // 1     '레코드구분
            string D002;  // 2     '자료구분
            string D003;  // 3     '세무서
            string D004;  // 6     '일련번호
            // 원천징수의무자
            string D005;  // 10    '사업자번호
            string D006;  // 50    '공란
            // 소득자
            string D007;  // 13    '소득자주민등록번호
            // 연금계좌입금명세
            string D008;  // 2     '연금계좌_일련번호
            string D009;  // 30    '연금계좌_취급자
            string D010;  // 10    '연금계좌_사업자등록번호
            string D011;  // 20    '연금계좌_계좌번호
            string D012;  // 8     '연금계좌_입금일
            string D013;  // 11    '연금계좌_계좌입금금액
            string D014;  // 944   '공란

            try
            {
                // D_RECORE QUERY
                sQry = "EXEC PH_PY995_D '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    while (!oRecordSet.EoF)
                    {
                        // D RECORD MOVE
                        D001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D001").Value.ToString().Trim(), 1);
                        D002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D002").Value.ToString().Trim(), 2);
                        D003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D003").Value.ToString().Trim(), 3);
                        D004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0');   // C레코드의 일련번호
                        D005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D005").Value.ToString().Trim(), 10);
                        D006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D006").Value.ToString().Trim(), 50);
                        D007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D007").Value.ToString().Trim(), 13);
                        D008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D008").Value.ToString().Trim(), 2);
                        D009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D009").Value.ToString().Trim(), 30);
                        D010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D010").Value.ToString().Trim(), 10);
                        D011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D011").Value.ToString().Trim(), 20);
                        D012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D012").Value.ToString().Trim(), 8, '0');
                        D013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D013").Value.ToString().Trim(), 11, '0');
                        D014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D014").Value.ToString().Trim(), 944);

                        FileSystem.PrintLine(1, D001 + D002 + D003 + D004 + D005 + D006 + D007 + D008 + D009 + D010 + D011 + D012 + D013 + D014);

                        oRecordSet.MoveNext();
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 필수 입력값 체크
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short errNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("HtaxID").Specific.VALUE))
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TeamName").Specific.VALUE))
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Dname").Specific.VALUE))
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Dtel").Specific.VALUE))
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE))
                {
                    errNum = 5;
                    throw new Exception();
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("홈텍스ID(5자리이상)를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("담당자부서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("담당자성명은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("담당자전화번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (File_Create() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            BubbleEvent = false;
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                            //사업장 변경되면
                            case "CLTCOD":
                                sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("HtaxID").Specific.String = oRecordSet.Fields.Item("U_HomeTId").Value.Trim();
                                oForm.Items.Item("TeamName").Specific.String = oRecordSet.Fields.Item("U_ChgDpt").Value.Trim();
                                oForm.Items.Item("Dname").Specific.String = oRecordSet.Fields.Item("U_ChgName").Value.Trim();
                                oForm.Items.Item("Dtel").Specific.String = oRecordSet.Fields.Item("U_ChgTel").Value.Trim();
                                oForm.ActiveItem = "DocDate"; //제출일자로 포커싱
                                break;

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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                BubbleEvent = false;
            }
            finally
            {
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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

                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1287": //복제
                            break;
                        case "1281":
                        case "1282":
                            oForm.Items.Item("JsnYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1288": // TODO: to "1291"
                            break;
                        case "1293":
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        #region Raise_FormItemEvent
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

        //					} else {
        //					}
        //					break;

        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					if (pval.BeforeAction == true) {

        //					} else if (pval.BeforeAction == false) {

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

        //					}
        //					break;
        //			}

        //			return;
        //			Raise_FormItemEvent_Error:
        //			///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Globals.Sbo_Application.StatusBar.SetText("Raise_FormItemEvent_Error:" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //		}
        #endregion

        #region Raise_FormMenuEvent
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{

        //			if (pval.BeforeAction == true) {
        //				return;
        //			}


        //			return;
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			int i = 0;
        //			string sQry = null;
        //			SAPbouiCOM.ComboBox oCombo = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;


        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			if ((BusinessObjectInfo.BeforeAction == false)) {

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
        #endregion

        #region 백업 소스코드
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="MatrixMsg"></param>
        ///// <param name="Insert_YN"></param>
        ///// <param name="MatrixErr"></param>
        //private void Matrix_AddRow(string MatrixMsg, bool Insert_YN = false, bool MatrixErr = false)
        //{
        //    //매트릭스 없는 Form인데 필요없는 메소드로 판단됨, 주석처리(2019.09.05 송명규)
        //    try
        //    {
        //        if (MatrixErr == true)
        //        {
        //            oForm.DataSources.UserDataSources.Item("Col0").Value = "??";
        //        }
        //        else
        //        {
        //            oForm.DataSources.UserDataSources.Item("Col0").Value = "";
        //        }
        //        oForm.DataSources.UserDataSources.Item("Col1").Value = MatrixMsg;
        //        if (Insert_YN == true)
        //        {
        //            oMat1.AddRow();
        //            MaxRow = MaxRow + 1;
        //        }
        //        oMat1.SetLineData(MaxRow);
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}
        #endregion
    }
}



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
//	internal class PH_PY995
//	{
//////  SAP MANAGE UI API 2004 SDK Sample
//////****************************************************************************
//////  File           : PH_PY995.cls
//////  Module         : 인사관리>정산관리
//////  Desc           : 퇴직소득지급명세서자료 전산매체수록
//////  FormType       :
//////  Create Date    : 2014.02.04
//////  Modified Date  : 2017.01.25
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


////2013년기준 1400 BYTE
////2014년기준 1090 BYTE
////2016년기준 1110 BYTE
////2017년기준 1110 BYTE

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
//				//세무대리인관리번호
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
//			[VBFixedString(932), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 932)]
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
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//대표자(성명)
//			public char[] B007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민(법인)등록번호
//			public char[] B008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//제출대상기간코드
//			public char[] B009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//퇴직소득자(C레코드)수
//			public char[] B010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//				//공란
//			public char[] B011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
//				//정산-과세대상퇴직금여합계
//			public char[] B012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] B013_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//신고대상소득세합계
//			public char[] B013_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//이연퇴직소득세액합계
//			public char[] B014;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] B015_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//차감원천징수-소득세액합계
//			public char[] B015_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] B016_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//차감원천징수-지방소득세액합계
//			public char[] B016_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] B017_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//차감원천징수-농어촌특별세액합계
//			public char[] B017_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] B018_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//차감원천징수세액-계합계
//			public char[] B018_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(893), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 893)]
//				//공란
//			public char[] B019;
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
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//징수의무자구분 1.사업장 3.공적연금사업자
//			public char[] C006;
//			///소득자
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//거주구분 1.거주자 2.비거주자
//			public char[] C007;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//내.외국인구분 1.내국인 9.외국인
//			public char[] C008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//거주지국코드
//			public char[] C009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//성명
//			public char[] C010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//주민등록번호
//			public char[] C011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//임원여부  1.여 2.부
//			public char[] C012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//확정급여형퇴직연금제도가입일
//			public char[] C013;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//2011.12.31 퇴직금
//			public char[] C014;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//귀속연도시작연월일
//			public char[] C015;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//귀속연도종료연월일
//			public char[] C016;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//퇴직사유
//			public char[] C017;
//			///퇴직급여현황-중간지급등
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//근무처명
//			public char[] C018;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//근무처사업자등록번호
//			public char[] C019;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직급여
//			public char[] C020;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//비과세퇴직급여
//			public char[] C021;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세대상퇴직급여
//			public char[] C022;
//			///퇴직급여현황-최종분
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(40), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 40)]
//				//근무처명
//			public char[] C023;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//근무처사업자등록번호
//			public char[] C024;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직급여
//			public char[] C025;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//비과세퇴직급여
//			public char[] C026;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세대상퇴직급여
//			public char[] C027;
//			///퇴직급여현황-정산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직급여
//			public char[] C028;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//비과세퇴직급여
//			public char[] C029;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세대상퇴직급여
//			public char[] C030;
//			///근속연수-중간지급등
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일
//			public char[] C031;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//기산일
//			public char[] C032;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일
//			public char[] C033;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//지급일
//			public char[] C034;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속월수
//			public char[] C035;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//제외월수
//			public char[] C036;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//가산월수
//			public char[] C037;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//중복월수
//			public char[] C038;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속연수
//			public char[] C039;
//			///근속연수-최종
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일
//			public char[] C040;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//기산일
//			public char[] C041;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일
//			public char[] C042;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//지급일
//			public char[] C043;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속월수
//			public char[] C044;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//제외월수
//			public char[] C045;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//가산월수
//			public char[] C046;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//중복월수
//			public char[] C047;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속연수
//			public char[] C048;
//			///근속연수-정산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일
//			public char[] C049;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//기산일
//			public char[] C050;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일
//			public char[] C051;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//지급일
//			public char[] C052;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속월수
//			public char[] C053;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//제외월수
//			public char[] C054;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//가산월수
//			public char[] C055;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//중복월수
//			public char[] C056;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속연수
//			public char[] C057;
//			///근속연수-안분-2012.12.31이전
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일
//			public char[] C058;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//기산일
//			public char[] C059;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일
//			public char[] C060;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//지급일
//			public char[] C061;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속월수
//			public char[] C062;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//제외월수
//			public char[] C063;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//가산월수
//			public char[] C064;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//중복월수
//			public char[] C065;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속연수
//			public char[] C066;
//			///근속연수-안분-2013.01.01이후
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일
//			public char[] C067;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//기산일
//			public char[] C068;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일
//			public char[] C069;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//지급일
//			public char[] C070;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속월수
//			public char[] C071;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//제외월수
//			public char[] C072;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//가산월수
//			public char[] C073;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//중복월수
//			public char[] C074;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//근속연수
//			public char[] C075;
//			///개정규정에따른계산방법-과세표준계산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득
//			public char[] C076;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//근속연수공제
//			public char[] C077;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산급여
//			public char[] C078;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산급여별공제
//			public char[] C079;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득과세표준
//			public char[] C080;
//			///개정규정에따른계산방법-세액계산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산산출세액
//			public char[] C081;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//산출세액
//			public char[] C082;
//			///종전규정에 따른 계산 방법 - 과세표준계산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득
//			public char[] C083;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득정률공제
//			public char[] C084;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//근속연수공제
//			public char[] C085;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득과세표준
//			public char[] C086;
//			///종전규정에 따른 계산 방법 - 세액계산 - 2012.12.31 이전
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세표준안분
//			public char[] C087;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균과세표준
//			public char[] C088;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산과세표준
//			public char[] C089;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산산출세액
//			public char[] C090;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균산출세액
//			public char[] C091;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//산출세액
//			public char[] C092;
//			///종전규정에 따른 계산 방법 - 세액계산 - 2013.1.1 이후
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세표준안분
//			public char[] C093;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균과세표준
//			public char[] C094;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산과세표준
//			public char[] C095;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산산출세액
//			public char[] C096;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균산출세액
//			public char[] C097;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//산출세액
//			public char[] C098;
//			///종전규정에 따른 계산 방법 - 세액계산 - 합계
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//과세표준안분
//			public char[] C099;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균과세표준
//			public char[] C100;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산과세표준
//			public char[] C101;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//환산산출세액
//			public char[] C102;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연평균산출세액
//			public char[] C103;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//산출세액
//			public char[] C104;
//			///퇴직소득세액계산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//				//퇴직일이속하는과세연도
//			public char[] C105;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직소득세산출세액
//			public char[] C106;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//기납부(또는기과세이연)세액
//			public char[] C107;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C108_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//신고대상세액
//			public char[] C108_2;
//			///이연퇴직소득세액계산
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C109_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//신고대상세액
//			public char[] C109_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//계좌입금금액_합계
//			public char[] C110;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//퇴직급여
//			public char[] C111;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//이연퇴직소득세
//			public char[] C112;
//			///납부명세-신고대상세액
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C113_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//소득세
//			public char[] C113_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C114_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//지방소득세
//			public char[] C114_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C115_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//농어총특별세
//			public char[] C115_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C116_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//계
//			public char[] C116_2;
//			///납부명세-이연퇴직소득세
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//소득세
//			public char[] C117;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//지방소득세
//			public char[] C118;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//농어촌특별세
//			public char[] C119;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//계
//			public char[] C120;
//			///납부명세-차감원천징수세액
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C121_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//소득세
//			public char[] C121_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C122_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//지방소득세
//			public char[] C122_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C123_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//농어총특별세
//			public char[] C123_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//				//부호
//			public char[] C124_1;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//계
//			public char[] C124_2;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//공란
//			public char[] C125;
//		}

//		C_record C_rec;

////2014년 D레코드 추가

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
//			///원천징수의무자
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//사업자번호
//			public char[] D005;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(50), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 50)]
//				//공란
//			public char[] D006;
//			///소득자
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(13), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//				//소득자주민등록번호
//			public char[] D007;
//			///연금계좌입금명세
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//				//연금계좌_일련번호
//			public char[] D008;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//				//연금계좌_취급자
//			public char[] D009;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//연금계좌_사업자등록번호
//			public char[] D010;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//연금계좌_계좌번호
//			public char[] D011;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//연금계좌_입금일
//			public char[] D012;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//				//연금계좌_계좌입금금액
//			public char[] D013;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(944), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 944)]
//				//공란
//			public char[] D014;
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


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY995.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			////여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount() * 10);

//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "PH_PY995_" + GetTotalFormsCount();

//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			//컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//			//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY995");
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
//						if (HeaderSpaceLineDel() == false) {
//							BubbleEvent = false;
//							return;
//						}
//						if (pval.ItemUID == "Btn01") {
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
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
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
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = Strings.Trim(oForm.Items.Item("YYYY").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			HtaxID = Strings.Trim(oForm.Items.Item("HtaxID").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TeamName = Strings.Trim(oForm.Items.Item("TeamName").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Dname = Strings.Trim(oForm.Items.Item("Dname").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Dtel = Strings.Trim(oForm.Items.Item("Dtel").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.VALUE);

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

//			/// C RECORD 처리
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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드(퇴직소득자 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//			sQry = "EXEC PH_PY995_A '" + CLTCOD + "', '" + HtaxID + "', '" + TeamName + "', '" + Dname + "', '" + Dtel + "', '" + DocDate + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				// PATH및 파일이름 만들기
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				saup = oRecordSet.Fields.Item("A009").Value;
//				//사업자번호
//				oFilePath = "C:\\BANK\\EA" + Strings.Mid(saup, 1, 7) + "." + Strings.Mid(saup, 8, 3);


//				//A RECORD MOVE

//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A001 = oRecordSet.Fields.Item("A001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A002 = oRecordSet.Fields.Item("A002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A003 = oRecordSet.Fields.Item("A003").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A004 = oRecordSet.Fields.Item("A004").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A005 = oRecordSet.Fields.Item("A005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A006 = oRecordSet.Fields.Item("A006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A007 = oRecordSet.Fields.Item("A007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A008 = oRecordSet.Fields.Item("A008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A009 = oRecordSet.Fields.Item("A009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A010 = oRecordSet.Fields.Item("A010").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A011 = oRecordSet.Fields.Item("A011").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A012 = oRecordSet.Fields.Item("A012").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A013 = oRecordSet.Fields.Item("A013").Value;

//				A_rec.A014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("A014").Value, new string("0", Strings.Len(A_rec.A014)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				A_rec.A015 = oRecordSet.Fields.Item("A015").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
//			sQry = "EXEC PH_PY995_B '" + CLTCOD + "', '" + yyyy + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			} else {
//				//B RECORD MOVE

//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B001 = oRecordSet.Fields.Item("B001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B002 = oRecordSet.Fields.Item("B002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B003 = oRecordSet.Fields.Item("B003").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B004 = oRecordSet.Fields.Item("B004").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B005 = oRecordSet.Fields.Item("B005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B006 = oRecordSet.Fields.Item("B006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B007 = oRecordSet.Fields.Item("B007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B008 = oRecordSet.Fields.Item("B008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B009 = oRecordSet.Fields.Item("B009").Value;
//				B_rec.B010 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B010").Value, new string("0", Strings.Len(B_rec.B010)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B011 = oRecordSet.Fields.Item("B011").Value;
//				B_rec.B012 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B012").Value, new string("0", Strings.Len(B_rec.B012)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B013_1 = oRecordSet.Fields.Item("B013_1").Value;
//				B_rec.B013_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B013_2").Value, new string("0", Strings.Len(B_rec.B013_2)));
//				B_rec.B014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B014").Value, new string("0", Strings.Len(B_rec.B014)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B015_1 = oRecordSet.Fields.Item("B015_1").Value;
//				B_rec.B015_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B015_2").Value, new string("0", Strings.Len(B_rec.B015_2)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B016_1 = oRecordSet.Fields.Item("B016_1").Value;
//				B_rec.B016_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B016_2").Value, new string("0", Strings.Len(B_rec.B016_2)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B017_1 = oRecordSet.Fields.Item("B017_1").Value;
//				B_rec.B017_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B017_2").Value, new string("0", Strings.Len(B_rec.B017_2)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B018_1 = oRecordSet.Fields.Item("B018_1").Value;
//				B_rec.B018_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("B018_2").Value, new string("0", Strings.Len(B_rec.B018_2)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				B_rec.B019 = oRecordSet.Fields.Item("B019").Value;

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref B_rec.B001) + MDC_SetMod.sStr(ref B_rec.B002) + MDC_SetMod.sStr(ref B_rec.B003) + MDC_SetMod.sStr(ref B_rec.B004) + MDC_SetMod.sStr(ref B_rec.B005) + MDC_SetMod.sStr(ref B_rec.B006) + MDC_SetMod.sStr(ref B_rec.B007) + MDC_SetMod.sStr(ref B_rec.B008) + MDC_SetMod.sStr(ref B_rec.B009) + MDC_SetMod.sStr(ref B_rec.B010) + MDC_SetMod.sStr(ref B_rec.B011) + MDC_SetMod.sStr(ref B_rec.B012) + MDC_SetMod.sStr(ref B_rec.B013_1) + MDC_SetMod.sStr(ref B_rec.B013_2) + MDC_SetMod.sStr(ref B_rec.B014) + MDC_SetMod.sStr(ref B_rec.B015_1) + MDC_SetMod.sStr(ref B_rec.B015_2) + MDC_SetMod.sStr(ref B_rec.B016_1) + MDC_SetMod.sStr(ref B_rec.B016_2) + MDC_SetMod.sStr(ref B_rec.B017_1) + MDC_SetMod.sStr(ref B_rec.B017_2) + MDC_SetMod.sStr(ref B_rec.B018_1) + MDC_SetMod.sStr(ref B_rec.B018_2) + MDC_SetMod.sStr(ref B_rec.B019));

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
//			sQry = "EXEC PH_PY995_C '" + CLTCOD + "', '" + yyyy + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, false);

//			NEWCNT = 0;

//			while (!(oRecordSet.EoF)) {

//				//C RECORD MOVE

//				NEWCNT = NEWCNT + 1;
//				/// 일련번호

//				//--D RECORD 호출용----------
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_SAUP = oRecordSet.Fields.Item("saup").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_YYYY = oRecordSet.Fields.Item("yyyy").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_SABUN = oRecordSet.Fields.Item("sabun").Value;
//				//--------------

//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C001 = oRecordSet.Fields.Item("C001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C002 = oRecordSet.Fields.Item("C002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C003 = oRecordSet.Fields.Item("C003").Value;
//				C_rec.C004 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(NEWCNT, new string("0", Strings.Len(C_rec.C004)));
//				/// 일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C005 = oRecordSet.Fields.Item("C005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C006 = oRecordSet.Fields.Item("C006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C007 = oRecordSet.Fields.Item("C007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C008 = oRecordSet.Fields.Item("C008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C009 = oRecordSet.Fields.Item("C009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C010 = oRecordSet.Fields.Item("C010").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C011 = oRecordSet.Fields.Item("C011").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C012 = oRecordSet.Fields.Item("C012").Value;
//				C_rec.C013 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C013").Value, new string("0", Strings.Len(C_rec.C013)));
//				C_rec.C014 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C014").Value, new string("0", Strings.Len(C_rec.C014)));
//				C_rec.C015 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C015").Value, new string("0", Strings.Len(C_rec.C015)));
//				C_rec.C016 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C016").Value, new string("0", Strings.Len(C_rec.C016)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C017 = oRecordSet.Fields.Item("C017").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C018 = oRecordSet.Fields.Item("C018").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C019 = oRecordSet.Fields.Item("C019").Value;
//				C_rec.C020 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C020").Value, new string("0", Strings.Len(C_rec.C020)));
//				C_rec.C021 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C021").Value, new string("0", Strings.Len(C_rec.C021)));
//				C_rec.C022 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C022").Value, new string("0", Strings.Len(C_rec.C022)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C023 = oRecordSet.Fields.Item("C023").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C024 = oRecordSet.Fields.Item("C024").Value;
//				C_rec.C025 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C025").Value, new string("0", Strings.Len(C_rec.C025)));
//				C_rec.C026 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C026").Value, new string("0", Strings.Len(C_rec.C026)));
//				C_rec.C027 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C027").Value, new string("0", Strings.Len(C_rec.C027)));
//				C_rec.C028 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C028").Value, new string("0", Strings.Len(C_rec.C028)));
//				C_rec.C029 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C029").Value, new string("0", Strings.Len(C_rec.C029)));
//				C_rec.C030 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C030").Value, new string("0", Strings.Len(C_rec.C030)));
//				C_rec.C031 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C031").Value, new string("0", Strings.Len(C_rec.C031)));
//				C_rec.C032 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C032").Value, new string("0", Strings.Len(C_rec.C032)));
//				C_rec.C033 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C033").Value, new string("0", Strings.Len(C_rec.C033)));
//				C_rec.C034 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C034").Value, new string("0", Strings.Len(C_rec.C034)));
//				C_rec.C035 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C035").Value, new string("0", Strings.Len(C_rec.C035)));
//				C_rec.C036 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C036").Value, new string("0", Strings.Len(C_rec.C036)));
//				C_rec.C037 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C037").Value, new string("0", Strings.Len(C_rec.C037)));
//				C_rec.C038 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C038").Value, new string("0", Strings.Len(C_rec.C038)));
//				C_rec.C039 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C039").Value, new string("0", Strings.Len(C_rec.C039)));
//				C_rec.C040 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C040").Value, new string("0", Strings.Len(C_rec.C040)));
//				C_rec.C041 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C041").Value, new string("0", Strings.Len(C_rec.C041)));
//				C_rec.C042 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C042").Value, new string("0", Strings.Len(C_rec.C042)));
//				C_rec.C043 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C043").Value, new string("0", Strings.Len(C_rec.C043)));
//				C_rec.C044 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C044").Value, new string("0", Strings.Len(C_rec.C044)));
//				C_rec.C045 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C045").Value, new string("0", Strings.Len(C_rec.C045)));
//				C_rec.C046 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C046").Value, new string("0", Strings.Len(C_rec.C046)));
//				C_rec.C047 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C047").Value, new string("0", Strings.Len(C_rec.C047)));
//				C_rec.C048 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C048").Value, new string("0", Strings.Len(C_rec.C048)));
//				C_rec.C049 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C049").Value, new string("0", Strings.Len(C_rec.C049)));
//				C_rec.C050 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C050").Value, new string("0", Strings.Len(C_rec.C050)));
//				C_rec.C051 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C051").Value, new string("0", Strings.Len(C_rec.C051)));
//				C_rec.C052 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C052").Value, new string("0", Strings.Len(C_rec.C052)));
//				C_rec.C053 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C053").Value, new string("0", Strings.Len(C_rec.C053)));
//				C_rec.C054 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C054").Value, new string("0", Strings.Len(C_rec.C054)));
//				C_rec.C055 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C055").Value, new string("0", Strings.Len(C_rec.C055)));
//				C_rec.C056 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C056").Value, new string("0", Strings.Len(C_rec.C056)));
//				C_rec.C057 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C057").Value, new string("0", Strings.Len(C_rec.C057)));
//				C_rec.C058 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C058").Value, new string("0", Strings.Len(C_rec.C058)));
//				C_rec.C059 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C059").Value, new string("0", Strings.Len(C_rec.C059)));
//				C_rec.C060 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C060").Value, new string("0", Strings.Len(C_rec.C060)));
//				C_rec.C061 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C061").Value, new string("0", Strings.Len(C_rec.C061)));
//				C_rec.C062 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C062").Value, new string("0", Strings.Len(C_rec.C062)));
//				C_rec.C063 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C063").Value, new string("0", Strings.Len(C_rec.C063)));
//				C_rec.C064 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C064").Value, new string("0", Strings.Len(C_rec.C064)));
//				C_rec.C065 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C065").Value, new string("0", Strings.Len(C_rec.C065)));
//				C_rec.C066 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C066").Value, new string("0", Strings.Len(C_rec.C066)));
//				C_rec.C067 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C067").Value, new string("0", Strings.Len(C_rec.C067)));
//				C_rec.C068 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C068").Value, new string("0", Strings.Len(C_rec.C068)));
//				C_rec.C069 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C069").Value, new string("0", Strings.Len(C_rec.C069)));
//				C_rec.C070 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C070").Value, new string("0", Strings.Len(C_rec.C070)));
//				C_rec.C071 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C071").Value, new string("0", Strings.Len(C_rec.C071)));
//				C_rec.C072 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C072").Value, new string("0", Strings.Len(C_rec.C072)));
//				C_rec.C073 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C073").Value, new string("0", Strings.Len(C_rec.C073)));
//				C_rec.C074 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C074").Value, new string("0", Strings.Len(C_rec.C074)));
//				C_rec.C075 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C075").Value, new string("0", Strings.Len(C_rec.C075)));
//				C_rec.C076 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C076").Value, new string("0", Strings.Len(C_rec.C076)));
//				C_rec.C077 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C077").Value, new string("0", Strings.Len(C_rec.C077)));
//				C_rec.C078 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C078").Value, new string("0", Strings.Len(C_rec.C078)));
//				C_rec.C079 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C079").Value, new string("0", Strings.Len(C_rec.C079)));
//				C_rec.C080 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C080").Value, new string("0", Strings.Len(C_rec.C080)));
//				C_rec.C081 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C081").Value, new string("0", Strings.Len(C_rec.C081)));
//				C_rec.C082 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C082").Value, new string("0", Strings.Len(C_rec.C082)));
//				C_rec.C083 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C083").Value, new string("0", Strings.Len(C_rec.C083)));
//				C_rec.C084 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C084").Value, new string("0", Strings.Len(C_rec.C084)));
//				C_rec.C085 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C085").Value, new string("0", Strings.Len(C_rec.C085)));
//				C_rec.C086 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C086").Value, new string("0", Strings.Len(C_rec.C086)));
//				C_rec.C087 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C087").Value, new string("0", Strings.Len(C_rec.C087)));
//				C_rec.C088 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C088").Value, new string("0", Strings.Len(C_rec.C088)));
//				C_rec.C089 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C089").Value, new string("0", Strings.Len(C_rec.C089)));
//				C_rec.C090 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C090").Value, new string("0", Strings.Len(C_rec.C090)));
//				C_rec.C091 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C091").Value, new string("0", Strings.Len(C_rec.C091)));
//				C_rec.C092 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C092").Value, new string("0", Strings.Len(C_rec.C092)));
//				C_rec.C093 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C093").Value, new string("0", Strings.Len(C_rec.C093)));
//				C_rec.C094 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C094").Value, new string("0", Strings.Len(C_rec.C094)));
//				C_rec.C095 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C095").Value, new string("0", Strings.Len(C_rec.C095)));
//				C_rec.C096 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C096").Value, new string("0", Strings.Len(C_rec.C096)));
//				C_rec.C097 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C097").Value, new string("0", Strings.Len(C_rec.C097)));
//				C_rec.C098 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C098").Value, new string("0", Strings.Len(C_rec.C098)));
//				C_rec.C099 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C099").Value, new string("0", Strings.Len(C_rec.C099)));
//				C_rec.C100 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C100").Value, new string("0", Strings.Len(C_rec.C100)));
//				C_rec.C101 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C101").Value, new string("0", Strings.Len(C_rec.C101)));
//				C_rec.C102 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C102").Value, new string("0", Strings.Len(C_rec.C102)));
//				C_rec.C103 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C103").Value, new string("0", Strings.Len(C_rec.C103)));
//				C_rec.C104 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C104").Value, new string("0", Strings.Len(C_rec.C104)));
//				C_rec.C105 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C105").Value, new string("0", Strings.Len(C_rec.C105)));
//				C_rec.C106 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C106").Value, new string("0", Strings.Len(C_rec.C106)));
//				C_rec.C107 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C107").Value, new string("0", Strings.Len(C_rec.C107)));
//				C_rec.C108_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C108_1").Value, new string("0", Strings.Len(C_rec.C108_1)));
//				C_rec.C108_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C108_2").Value, new string("0", Strings.Len(C_rec.C108_2)));
//				C_rec.C109_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C109_1").Value, new string("0", Strings.Len(C_rec.C109_1)));
//				C_rec.C109_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C109_2").Value, new string("0", Strings.Len(C_rec.C109_2)));
//				C_rec.C110 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C110").Value, new string("0", Strings.Len(C_rec.C110)));
//				C_rec.C111 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C111").Value, new string("0", Strings.Len(C_rec.C111)));
//				C_rec.C112 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C112").Value, new string("0", Strings.Len(C_rec.C112)));
//				C_rec.C113_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C113_1").Value, new string("0", Strings.Len(C_rec.C113_1)));
//				C_rec.C113_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C113_2").Value, new string("0", Strings.Len(C_rec.C113_2)));
//				C_rec.C114_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C114_1").Value, new string("0", Strings.Len(C_rec.C114_1)));
//				C_rec.C114_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C114_2").Value, new string("0", Strings.Len(C_rec.C114_2)));
//				C_rec.C115_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C115_1").Value, new string("0", Strings.Len(C_rec.C115_1)));
//				C_rec.C115_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C115_2").Value, new string("0", Strings.Len(C_rec.C115_2)));
//				C_rec.C116_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C116_1").Value, new string("0", Strings.Len(C_rec.C116_1)));
//				C_rec.C116_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C116_2").Value, new string("0", Strings.Len(C_rec.C116_2)));
//				C_rec.C117 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C117").Value, new string("0", Strings.Len(C_rec.C117)));
//				C_rec.C118 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C118").Value, new string("0", Strings.Len(C_rec.C118)));
//				C_rec.C119 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C119").Value, new string("0", Strings.Len(C_rec.C119)));
//				C_rec.C120 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C120").Value, new string("0", Strings.Len(C_rec.C120)));
//				C_rec.C121_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C121_1").Value, new string("0", Strings.Len(C_rec.C121_1)));
//				C_rec.C121_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C121_2").Value, new string("0", Strings.Len(C_rec.C121_2)));
//				C_rec.C122_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C122_1").Value, new string("0", Strings.Len(C_rec.C122_1)));
//				C_rec.C122_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C122_2").Value, new string("0", Strings.Len(C_rec.C122_2)));
//				C_rec.C123_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C123_1").Value, new string("0", Strings.Len(C_rec.C123_1)));
//				C_rec.C123_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C123_2").Value, new string("0", Strings.Len(C_rec.C123_2)));
//				C_rec.C124_1 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C124_1").Value, new string("0", Strings.Len(C_rec.C124_1)));
//				C_rec.C124_2 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("C124_2").Value, new string("0", Strings.Len(C_rec.C124_2)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				C_rec.C125 = oRecordSet.Fields.Item("C125").Value;

//				//예제
//				//C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").VALUE, "-", "")

//				//OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").VALUE) + Val(oRecordSet.Fields("U_BIGWA3").VALUE) + Val(oRecordSet.Fields("U_BIGWA5").VALUE) _
//				//'        + Val(oRecordSet.Fields("U_BIGWA6").VALUE) + Val(oRecordSet.Fields("U_BIGWU3").VALUE)

//				//C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
//				//C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
//				//C_rec.FILLER = Space$(Len(C_rec.FILLER))
//				//C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, , String$(Len(C_rec.C022), "0"))


//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref C_rec.C001) + MDC_SetMod.sStr(ref C_rec.C002) + MDC_SetMod.sStr(ref C_rec.C003) + MDC_SetMod.sStr(ref C_rec.C004) + MDC_SetMod.sStr(ref C_rec.C005) + MDC_SetMod.sStr(ref C_rec.C006) + MDC_SetMod.sStr(ref C_rec.C007) + MDC_SetMod.sStr(ref C_rec.C008) + MDC_SetMod.sStr(ref C_rec.C009) + MDC_SetMod.sStr(ref C_rec.C010) + MDC_SetMod.sStr(ref C_rec.C011) + MDC_SetMod.sStr(ref C_rec.C012) + MDC_SetMod.sStr(ref C_rec.C013) + MDC_SetMod.sStr(ref C_rec.C014) + MDC_SetMod.sStr(ref C_rec.C015) + MDC_SetMod.sStr(ref C_rec.C016) + MDC_SetMod.sStr(ref C_rec.C017) + MDC_SetMod.sStr(ref C_rec.C018) + MDC_SetMod.sStr(ref C_rec.C019) + MDC_SetMod.sStr(ref C_rec.C020) + MDC_SetMod.sStr(ref C_rec.C021) + MDC_SetMod.sStr(ref C_rec.C022) + MDC_SetMod.sStr(ref C_rec.C023) + MDC_SetMod.sStr(ref C_rec.C024) + MDC_SetMod.sStr(ref C_rec.C025) + MDC_SetMod.sStr(ref C_rec.C026) + MDC_SetMod.sStr(ref C_rec.C027) + MDC_SetMod.sStr(ref C_rec.C028) + MDC_SetMod.sStr(ref C_rec.C029) + MDC_SetMod.sStr(ref C_rec.C030) + MDC_SetMod.sStr(ref C_rec.C031) + MDC_SetMod.sStr(ref C_rec.C032) + MDC_SetMod.sStr(ref C_rec.C033) + MDC_SetMod.sStr(ref C_rec.C034) + MDC_SetMod.sStr(ref C_rec.C035) + MDC_SetMod.sStr(ref C_rec.C036) + MDC_SetMod.sStr(ref C_rec.C037) + MDC_SetMod.sStr(ref C_rec.C038) + MDC_SetMod.sStr(ref C_rec.C039) + MDC_SetMod.sStr(ref C_rec.C040) + MDC_SetMod.sStr(ref C_rec.C041) + MDC_SetMod.sStr(ref C_rec.C042) + MDC_SetMod.sStr(ref C_rec.C043) + MDC_SetMod.sStr(ref C_rec.C044) + MDC_SetMod.sStr(ref C_rec.C045) + MDC_SetMod.sStr(ref C_rec.C046) + MDC_SetMod.sStr(ref C_rec.C047) + MDC_SetMod.sStr(ref C_rec.C048) + MDC_SetMod.sStr(ref C_rec.C049) + MDC_SetMod.sStr(ref C_rec.C050) + MDC_SetMod.sStr(ref C_rec.C051) + MDC_SetMod.sStr(ref C_rec.C052) + MDC_SetMod.sStr(ref C_rec.C053) + MDC_SetMod.sStr(ref C_rec.C054) + MDC_SetMod.sStr(ref C_rec.C055) + MDC_SetMod.sStr(ref C_rec.C056) + MDC_SetMod.sStr(ref C_rec.C057) + MDC_SetMod.sStr(ref C_rec.C058) + MDC_SetMod.sStr(ref C_rec.C059) + MDC_SetMod.sStr(ref C_rec.C060) + MDC_SetMod.sStr(ref C_rec.C061) + MDC_SetMod.sStr(ref C_rec.C062) + MDC_SetMod.sStr(ref C_rec.C063) + MDC_SetMod.sStr(ref C_rec.C064) + MDC_SetMod.sStr(ref C_rec.C065) + MDC_SetMod.sStr(ref C_rec.C066) + MDC_SetMod.sStr(ref C_rec.C067) + MDC_SetMod.sStr(ref C_rec.C068) + MDC_SetMod.sStr(ref C_rec.C069) + MDC_SetMod.sStr(ref C_rec.C070) + MDC_SetMod.sStr(ref C_rec.C071) + MDC_SetMod.sStr(ref C_rec.C072) + MDC_SetMod.sStr(ref C_rec.C073) + MDC_SetMod.sStr(ref C_rec.C074) + MDC_SetMod.sStr(ref C_rec.C075) + MDC_SetMod.sStr(ref C_rec.C076) + MDC_SetMod.sStr(ref C_rec.C077) + MDC_SetMod.sStr(ref C_rec.C078) + MDC_SetMod.sStr(ref C_rec.C079) + MDC_SetMod.sStr(ref C_rec.C080) + MDC_SetMod.sStr(ref C_rec.C081) + MDC_SetMod.sStr(ref C_rec.C082) + MDC_SetMod.sStr(ref C_rec.C083) + MDC_SetMod.sStr(ref C_rec.C084) + MDC_SetMod.sStr(ref C_rec.C085) + MDC_SetMod.sStr(ref C_rec.C086) + MDC_SetMod.sStr(ref C_rec.C087) + MDC_SetMod.sStr(ref C_rec.C088) + MDC_SetMod.sStr(ref C_rec.C089) + MDC_SetMod.sStr(ref C_rec.C090) + MDC_SetMod.sStr(ref C_rec.C091) + MDC_SetMod.sStr(ref C_rec.C092) + MDC_SetMod.sStr(ref C_rec.C093) + MDC_SetMod.sStr(ref C_rec.C094) + MDC_SetMod.sStr(ref C_rec.C095) + MDC_SetMod.sStr(ref C_rec.C096) + MDC_SetMod.sStr(ref C_rec.C097) + MDC_SetMod.sStr(ref C_rec.C098) + MDC_SetMod.sStr(ref C_rec.C099) + MDC_SetMod.sStr(ref C_rec.C100) + MDC_SetMod.sStr(ref C_rec.C101) + MDC_SetMod.sStr(ref C_rec.C102) + MDC_SetMod.sStr(ref C_rec.C103) + MDC_SetMod.sStr(ref C_rec.C104) + MDC_SetMod.sStr(ref C_rec.C105) + MDC_SetMod.sStr(ref C_rec.C106) + MDC_SetMod.sStr(ref C_rec.C107) + MDC_SetMod.sStr(ref C_rec.C108_1) + MDC_SetMod.sStr(ref C_rec.C108_2) + MDC_SetMod.sStr(ref C_rec.C109_1) + MDC_SetMod.sStr(ref C_rec.C109_2) + MDC_SetMod.sStr(ref C_rec.C110) + MDC_SetMod.sStr(ref C_rec.C111) + MDC_SetMod.sStr(ref C_rec.C112) + MDC_SetMod.sStr(ref C_rec.C113_1) + MDC_SetMod.sStr(ref C_rec.C113_2) + MDC_SetMod.sStr(ref C_rec.C114_1) + MDC_SetMod.sStr(ref C_rec.C114_2) + MDC_SetMod.sStr(ref C_rec.C115_1) + MDC_SetMod.sStr(ref C_rec.C115_2) + MDC_SetMod.sStr(ref C_rec.C116_1) + MDC_SetMod.sStr(ref C_rec.C116_2) + MDC_SetMod.sStr(ref C_rec.C117) + MDC_SetMod.sStr(ref C_rec.C118) + MDC_SetMod.sStr(ref C_rec.C119) + MDC_SetMod.sStr(ref C_rec.C120) + MDC_SetMod.sStr(ref C_rec.C121_1) + MDC_SetMod.sStr(ref C_rec.C121_2) + MDC_SetMod.sStr(ref C_rec.C122_1) + MDC_SetMod.sStr(ref C_rec.C122_2) + MDC_SetMod.sStr(ref C_rec.C123_1) + MDC_SetMod.sStr(ref C_rec.C123_2) + MDC_SetMod.sStr(ref C_rec.C124_1) + MDC_SetMod.sStr(ref C_rec.C124_2) + MDC_SetMod.sStr(ref C_rec.C125));


//				/// D레코드: 연금계좌입금명세 레코드 수록
//				///계좌입금금액_합계  2016부터있음
//				if (Conversion.Val(C_rec.C110) > 0) {
//					if (File_Create_D_record() == false) {
//						ErrNum = 2;
//						goto Error_Message;
//					}
//				}


//				oRecordSet.MoveNext();

//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";


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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//			short JONCNT = 0;

//			CheckD = Convert.ToString(false);
//			///체크필요유무
//			ErrNum = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// D_RECORE QUERY
//			sQry = "EXEC PH_PY995_D '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto Error_Message;
//			}

//			while (!(oRecordSet.EoF)) {

//				//D RECORD MOVE

//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D001 = oRecordSet.Fields.Item("D001").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D002 = oRecordSet.Fields.Item("D002").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D003 = oRecordSet.Fields.Item("D003").Value;
//				D_rec.D004 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(C_rec.C004, new string("0", Strings.Len(D_rec.D004)));
//				/// C레코드의 일련번호
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D005 = oRecordSet.Fields.Item("D005").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D006 = oRecordSet.Fields.Item("D006").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D007 = oRecordSet.Fields.Item("D007").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D008 = oRecordSet.Fields.Item("D008").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D009 = oRecordSet.Fields.Item("D009").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D010 = oRecordSet.Fields.Item("D010").Value;
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D011 = oRecordSet.Fields.Item("D011").Value;
//				D_rec.D012 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D012").Value, new string("0", Strings.Len(D_rec.D012)));
//				D_rec.D013 = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("D013").Value, new string("0", Strings.Len(D_rec.D013)));
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				D_rec.D014 = oRecordSet.Fields.Item("D014").Value;

//				FileSystem.PrintLine(1, MDC_SetMod.sStr(ref D_rec.D001) + MDC_SetMod.sStr(ref D_rec.D002) + MDC_SetMod.sStr(ref D_rec.D003) + MDC_SetMod.sStr(ref D_rec.D004) + MDC_SetMod.sStr(ref D_rec.D005) + MDC_SetMod.sStr(ref D_rec.D006) + MDC_SetMod.sStr(ref D_rec.D007) + MDC_SetMod.sStr(ref D_rec.D008) + MDC_SetMod.sStr(ref D_rec.D009) + MDC_SetMod.sStr(ref D_rec.D010) + MDC_SetMod.sStr(ref D_rec.D011) + MDC_SetMod.sStr(ref D_rec.D012) + MDC_SetMod.sStr(ref D_rec.D013) + MDC_SetMod.sStr(ref D_rec.D014));

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
//				MDC_Globals.Sbo_Application.StatusBar.SetText("D레코드가 존재하지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
//			//UPGRADE_WARNING: oForm.Items(HtaxID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(oForm.Items.Item("HtaxID").Specific.VALUE)) {
//				ErrNum = 1;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(TeamName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("TeamName").Specific.VALUE)) {
//				ErrNum = 2;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(Dname).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("Dname").Specific.VALUE)) {
//				ErrNum = 3;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(Dtel).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("Dtel").Specific.VALUE)) {
//				ErrNum = 4;
//				goto HeaderSpaceLineDel;
//				//UPGRADE_WARNING: oForm.Items(DocDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			} else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE)) {
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
