using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 퇴직소득지급명세서자료 전산매체수록
    /// </summary>
    internal class PH_PY995 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        
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
            string CLTCOD = string.Empty;
            string yyyy = string.Empty;
            string HtaxID = string.Empty;
            string TeamName = string.Empty;
            string Dname = string.Empty;
            string Dtel = string.Empty;
            string DocDate = string.Empty;

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
                HtaxID = oForm.Items.Item("HtaxID").Specific.Value.ToString().Trim();
                TeamName = oForm.Items.Item("TeamName").Specific.Value.ToString().Trim();
                Dname = oForm.Items.Item("Dname").Specific.Value.ToString().Trim();
                Dtel = oForm.Items.Item("Dtel").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

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
            // 2020년기준 761 BYTE

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
            string A016;  //  583   '공란

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
                    A016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A016").Value.ToString().Trim(), 583);

                    FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015 + A016);
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
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
            string B019;    // 544   '공란

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
                    B001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B001").Value.ToString().Trim(), 1);
                    B002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B002").Value.ToString().Trim(), 2);
                    B003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B003").Value.ToString().Trim(), 3);
                    B004 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B004").Value.ToString().Trim(), 6);
                    B005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B005").Value.ToString().Trim(), 10);
                    B006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B006").Value.ToString().Trim(), 40);
                    B007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B007").Value.ToString().Trim(), 30);
                    B008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B008").Value.ToString().Trim(), 13);
                    B009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B009").Value.ToString().Trim(), 1);
                    B010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B010").Value.ToString().Trim(), 7, '0');
                    B011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B011").Value.ToString().Trim(), 7);
                    B012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B012").Value.ToString().Trim(), 14, '0');
                    B013_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013_1").Value.ToString().Trim(), 1);
                    B013_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013_2").Value.ToString().Trim(), 13, '0');
                    B014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B014").Value.ToString().Trim(), 13, '0');
                    B015_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B015_1").Value.ToString().Trim(), 1);
                    B015_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B015_2").Value.ToString().Trim(), 13, '0');
                    B016_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B016_1").Value.ToString().Trim(), 1);
                    B016_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B016_2").Value.ToString().Trim(), 13, '0');
                    B017_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B017_1").Value.ToString().Trim(), 1);
                    B017_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B017_2").Value.ToString().Trim(), 13, '0');
                    B018_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018_1").Value.ToString().Trim(), 1);
                    B018_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018_2").Value.ToString().Trim(), 13, '0');
                    B019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B019").Value.ToString().Trim(), 544);

                    FileSystem.PrintLine(1, B001 + B002 + B003 + B004 + B005 + B006 + B007 + B008 + B009 + B010 + B011 + B012 + B013_1 + B013_2 + B014
                                          + B015_1 + B015_2 + B016_1 + B016_2 + B017_1 + B017_2 + B018_1 + B018_2 + B019);
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
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
            string C_SAUP = string.Empty;
            string C_YYYY = string.Empty;
            string C_SABUN = string.Empty;
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
            string C009;   // 1     '종교관련종사자여부
            string C010;   // 2     '거주지국코드
            string C011;   // 30    '성명
            string C012;   // 13    '주민등록번호
            string C013;   // 1     '임원여부  1.여 2.부
            string C014;   // 8     '확정급여형퇴직연금제도가입일
            string C015;   // 11    '2011.12.31 퇴직금
            string C016;   // 8     '귀속연도시작연월일
            string C017;   // 8     '귀속연도종료연월일
            string C018;   // 1     '퇴직사유
            // 퇴직급여현황-중간지급등
            string C019;   // 40    '근무처명
            string C020;   // 10    '근무처사업자등록번호
            string C021;   // 11    '퇴직급여
            string C022;   // 11    '비과세퇴직급여
            string C023;   // 11    '과세대상퇴직급여
            // 퇴직급여현황-최종분
            string C024;   // 40    '근무처명
            string C025;   // 10    '근무처사업자등록번호
            string C026;   // 11    '퇴직급여
            string C027;   // 11    '비과세퇴직급여
            string C028;   // 11    '과세대상퇴직급여
            // 퇴직급여현황-정산
            string C029;   // 11    '퇴직급여
            string C030;   // 11    '비과세퇴직급여
            string C031;   // 11    '과세대상퇴직급여
            // 근속연수-중간지급등
            string C032;   // 8     '입사일
            string C033;   // 8     '기산일
            string C034;   // 8     '퇴사일
            string C035;   // 8     '지급일
            string C036;   // 4     '근속월수
            string C037;   // 4     '제외월수
            string C038;   // 4     '가산월수
            string C039;   // 4     '중복월수
            string C040;   // 4     '근속연수
            // 근속연수-최종
            string C041;   // 8     '입사일
            string C042;   // 8     '기산일
            string C043;   // 8     '퇴사일
            string C044;   // 8     '지급일
            string C045;   // 4     '근속월수
            string C046;   // 4     '제외월수
            string C047;   // 4     '가산월수
            string C048;   // 4     '중복월수
            string C049;   // 4     '근속연수
            // 근속연수-정산
            string C050;   // 8     '입사일
            string C051;   // 8     '기산일
            string C052;   // 8     '퇴사일
            string C053;   // 8     '지급일
            string C054;   // 4     '근속월수
            string C055;   // 4     '제외월수
            string C056;   // 4     '가산월수
            string C057;   // 4     '중복월수
            string C058;   // 4     '근속연수
            // 과세표준계산
            string C059;   // 11    '퇴직소득
            string C060;   // 11    '근속연수공제
            string C061;   // 11    '환산급여
            string C062;   // 11    '환산급여별공제
            string C063;   // 11    '퇴직소득과세표준
            // 퇴직소득세액계산
            string C064;   // 11    '환산산출세액
            string C065;   // 11    '산출세액
            string C066;   // 11    '퇴직소득세산출세액
            string C067;   // 11    '기납부(또는기과세이연)세액
            string C068_1; // 1     '부호
            string C068_2; // 11    '신고대상세액
            // 이연퇴직소득세액계산
            string C069_1; // 1     '부호
            string C069_2; // 11    '신고대상세액
            string C070;   // 11    '계좌입금금액_합계
            string C071;   // 11    '퇴직급여
            string C072;   // 11    '이연퇴직소득세
            // 납부명세-신고대상세액
            string C073_1; // 1     '부호
            string C073_2; // 11    '소득세
            string C074_1; // 1     '부호
            string C074_2; // 11    '지방소득세
            string C075_1; // 1     '부호
            string C075_2; // 11    '농어총특별세
            string C076_1; // 1     '부호
            string C076_2; // 11    '계
            // 납부명세-이연퇴직소득세
            string C077;   // 11    '소득세
            string C078;   // 11    '지방소득세
            string C079;   // 11    '농어촌특별세
            string C080;   // 11    '계
            // 납부명세-차감원천징수세액
            string C081_1; // 1     '부호
            string C081_2; // 11    '소득세
            string C082_1; // 1     '부호
            string C082_2; // 11    '지방소득세
            string C083_1; // 1     '부호
            string C083_2; // 11    '농어총특별세
            string C084_1; // 1     '부호
            string C084_2; // 11    '계
            string C085;   // 2     '공란

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
                        NEWCNT += 1; //일련번호 + 1
                        C_SAUP = oRecordSet.Fields.Item("saup").Value.ToString().Trim();
                        C_YYYY = oRecordSet.Fields.Item("yyyy").Value.ToString().Trim();
                        C_SABUN = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();

                        //C RECORD MOVE
                        C001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C001").Value.ToString().Trim(), 1);
                        C002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C002").Value.ToString().Trim(), 2);
                        C003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C003").Value.ToString().Trim(), 3);
                        C004 = codeHelpClass.GetFixedLengthStringByte(NEWCNT.ToString(), 6, '0');
                        C005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C005").Value.ToString().Trim(), 10);
                        C006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C006").Value.ToString().Trim(), 1, '0');
                        C007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C007").Value.ToString().Trim(), 1);
                        C008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C008").Value.ToString().Trim(), 1);
                        C009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C009").Value.ToString().Trim(), 1);
                        C010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C010").Value.ToString().Trim(), 2);

                        C011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C011").Value.ToString().Trim(), 30);
                        C012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C012").Value.ToString().Trim(), 13);
                        C013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C013").Value.ToString().Trim(), 1);
                        C014 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C014").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        // TABLE이 소수 6자리(numeric(19,6))라서 사사오입 시킴(C#은 기본으로 .5가 반올림 안됨 그래서 MidpointRounding.AwayFromZero 문을 씀)
                        C015 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C015").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C016 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C016").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        C017 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C017").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 8, '0');
                        C018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C018").Value.ToString().Trim(), 1);
                        C019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C019").Value.ToString().Trim(), 40);
                        C020 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C020").Value.ToString().Trim(), 10);

                        C021 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C021").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C022 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C022").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C023 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C023").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C024 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C024").Value.ToString().Trim(), 40);
                        C025 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C025").Value.ToString().Trim(), 10);
                        C026 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C026").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C027 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C027").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C028 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C028").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C029 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C029").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C030 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C030").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');

                        C031 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C031").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C032 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C032").Value.ToString().Trim(), 8, '0');
                        C033 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C033").Value.ToString().Trim(), 8, '0');
                        C034 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C034").Value.ToString().Trim(), 8, '0');
                        C035 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C035").Value.ToString().Trim(), 8, '0');
                        C036 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C036").Value.ToString().Trim(), 4, '0');
                        C037 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C037").Value.ToString().Trim(), 4, '0');
                        C038 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C038").Value.ToString().Trim(), 4, '0');
                        C039 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C039").Value.ToString().Trim(), 4, '0');
                        C040 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C040").Value.ToString().Trim(), 4, '0');

                        C041 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C041").Value.ToString().Trim(), 8, '0');
                        C042 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C042").Value.ToString().Trim(), 8, '0');
                        C043 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C043").Value.ToString().Trim(), 8, '0');
                        C044 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C044").Value.ToString().Trim(), 8, '0');
                        C045 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C045").Value.ToString().Trim(), 4, '0');
                        C046 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C046").Value.ToString().Trim(), 4, '0');
                        C047 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C047").Value.ToString().Trim(), 4, '0');
                        C048 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C048").Value.ToString().Trim(), 4, '0');
                        C049 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C049").Value.ToString().Trim(), 4, '0');
                        C050 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C050").Value.ToString().Trim(), 8, '0');

                        C051 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C051").Value.ToString().Trim(), 8, '0');
                        C052 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C052").Value.ToString().Trim(), 8, '0');
                        C053 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C053").Value.ToString().Trim(), 8, '0');
                        C054 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C054").Value.ToString().Trim(), 4, '0');
                        C055 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C055").Value.ToString().Trim(), 4, '0');
                        C056 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C056").Value.ToString().Trim(), 4, '0');
                        C057 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C057").Value.ToString().Trim(), 4, '0');
                        C058 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C058").Value.ToString().Trim(), 4, '0');
                        C059 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C059").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C060 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C060").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');

                        C061 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C061").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C062 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C062").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C063 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C063").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C064 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C064").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C065 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C065").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C066 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C066").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C067 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C067").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C068_1 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C068_1").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 1, '0');
                        C068_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C068_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C069_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C069_1").Value.ToString().Trim(), 1, '0');
                        C069_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C069_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C070 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C070").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');

                        C071 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C071").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C072 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C072").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C073_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C073_1").Value.ToString().Trim(), 1, '0');
                        C073_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C073_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C074_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C074_1").Value.ToString().Trim(), 1, '0');
                        C074_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C074_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C075_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C075_1").Value.ToString().Trim(), 1, '0');
                        C075_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C075_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C076_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C076_1").Value.ToString().Trim(), 1, '0');
                        C076_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C076_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C077 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C077").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C078 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C078").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C079 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C079").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C080 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C080").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');

                        C081_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C081_1").Value.ToString().Trim(), 1, '0');
                        C081_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C081_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C082_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C082_1").Value.ToString().Trim(), 1, '0');
                        C082_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C082_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C083_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C083_1").Value.ToString().Trim(), 1, '0');
                        C083_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C083_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C084_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C084_1").Value.ToString().Trim(), 1, '0');
                        C084_2 = codeHelpClass.GetFixedLengthStringByte(Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("C084_2").Value), MidpointRounding.AwayFromZero).ToString().Trim(), 11, '0');
                        C085 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C085").Value.ToString().Trim(), 2);

                        FileSystem.PrintLine(1, C001 + C002 + C003 + C004 + C005 + C006 + C007 + C008 + C009 + C010 + C011 + C012 + C013 + C014 + C015 + C016 + C017 + C018 + C019 + C020
                                              + C021 + C022 + C023 + C024 + C025 + C026 + C027 + C028 + C029 + C030 + C031 + C032 + C033 + C034 + C035 + C036 + C037 + C038 + C039 + C040
                                              + C041 + C042 + C043 + C044 + C045 + C046 + C047 + C048 + C049 + C050 + C051 + C052 + C053 + C054 + C055 + C056 + C057 + C058 + C059 + C060
                                              + C061 + C062 + C063 + C064 + C065 + C066 + C067 + C068_1 + C068_2 + C069_1 + C069_2 + C070 + C071 + C072 + C073_1 + C073_2 + C074_1 + C074_2 + C075_1 + C075_2 + C076_1 + C076_2 + C077 + C078 + C079 + C080
                                              + C081_1 + C081_2 + C082_1 + C082_2 + C083_1 + C083_2 + C084_1 + C084_2 + C085);

                        // D 레코드 : 연금계좌입금명세 레코드 수록
                        if (Conversion.Val(C070) > 0)    // 계좌입금금액_합계  2016부터있음
                        {
                            if (File_Create_D_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                            {
                                errNum = 2;
                                throw new Exception();
                            }
                        }

                        oRecordSet.MoveNext();

                        ProgressBar01.Value += 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
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
            string D014;  // 595   '공란

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
                        D014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D014").Value.ToString().Trim(), 595);

                        FileSystem.PrintLine(1, D001 + D002 + D003 + D004 + D005 + D006 + D007 + D008 + D009 + D010 + D011 + D012 + D013 + D014);

                        oRecordSet.MoveNext();
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
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
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:        //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:            //2
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:           //3
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:        //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK:               //6
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:        //7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE:            //10
                    break;                                           
                                                                     
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:         //11
                    break;                                           
                                                                     
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:         //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:         //21
                    break;                                           
                                                                     
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:    //27
                    break;
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
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
    }
}
 
