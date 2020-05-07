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
    /// 근로소득지급명세서자료 전산매체수록
    /// </summary>
    internal class PH_PY980 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY980.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY980_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY980");

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
                if (File_Create_A_record(CLTCOD, yyyy, HtaxID, TeamName, Dname, Dtel, DocDate) == false)
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
        private bool File_Create_A_record(string pCLTCOD, string pyyyy, string pHtaxID, string pTeamName, string pDname, string pDtel, string pDocDate)
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
            // 2014년기준 1520 BYTE
            // 2014년기준 1580 BYTE  re
            // 2015년귀속 1610 BYTE
            // 2016년귀속 1620 BYTE
            // 2017년귀속 1620 BYTE
            // 2018년귀속 1882 BYTE
            // 2019년귀속 2082 BYTE

            string A001; // 1     '레코드구분
            string A002; // 2     '자료구분
            string A003; // 3     '세무서코드
            string A004; // 8     '제출일자
            string A005; // 1     '제출자구분 (1;세무대리인, 2;법인, 3;개인)
            string A006; // 6     '세무대리인
            string A007; // 20    '홈텍스ID
            string A008; // 4     '세무프로그램코드
            string A009; // 10    '사업자번호
            string A010; // 60    '법인명(상호)
            string A011; // 30    '담당자부서
            string A012; // 30    '담당자성명
            string A013; // 15    '담당자전화번호
            string A014; // 4     '귀속년도          --2019
            string A015; // 5     '신고의무자수
            string A016; // 3     '한글코드종류
            string A017; // 1880  '공란

            try
            {
                //A_RECORE QUERY
                sQry = "EXEC PH_PY980_A '" + pCLTCOD + "', '" + pyyyy + "', '" + pHtaxID + "', '" + pTeamName + "', '" + pDname + "', '" + pDtel + "', '" + pDocDate + "'";
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
                    oFilePath = "C:\\BANK\\C" + codeHelpClass.Mid(saup, 0, 7) + "." + codeHelpClass.Mid(saup, 7, 3);
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
                    A010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A010").Value.ToString().Trim(), 60);
                    A011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A011").Value.ToString().Trim(), 30);
                    A012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A012").Value.ToString().Trim(), 30);
                    A013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A013").Value.ToString().Trim(), 15);
                    A014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A014").Value.ToString().Trim(), 4);
                    A015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A015").Value.ToString().Trim(), 5, '0');
                    A016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A016").Value.ToString().Trim(), 3);
                    A017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A017").Value.ToString().Trim(), 1880);

                    FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015 + A016 + A017);

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
            string B001; // 1     '레코드구분
            string B002; // 2     '자료구분
            string B003; // 3     '세무서
            string B004; // 6     '일련번호
            string B005; // 10    '사업자번호
            string B006; // 60    '법인명(상호)
            string B007; // 30    '대표자
            string B008; // 13    '주민(법인)번호
            string B009; // 4     '귀속년도             --2019
            string B010; // 7     '주(현)근무처(C레코드)수
            string B011; // 7     '종(전)근무처(D레코드)수
            string B012; // 14    '총급여총계
            string B013; // 13    '결정세액(소득세)총계
            string B014; // 13    '결정세액(지방소득세)총계
            string B015; // 13    '결정세액(농특세)총계
            string B016; // 13    '결정세액총계
            string B017; // 1     '제출대상기간
            string B018; // 1872  '공란

            try
            {
                // B_RECORE QUERY
                sQry = "EXEC PH_PY980_B '" + pCLTCOD + "', '" + pyyyy + "'";
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
                    B006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B006").Value.ToString().Trim(), 60);
                    B007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B007").Value.ToString().Trim(), 30);
                    B008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B008").Value.ToString().Trim(), 13);
                    B009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B009").Value.ToString().Trim(), 4);
                    B010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B010").Value.ToString().Trim(), 7, '0');
                    B011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B011").Value.ToString().Trim(), 7, '0');
                    B012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B012").Value.ToString().Trim(), 14, '0');
                    B013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013").Value.ToString().Trim(), 13, '0');
                    B014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B014").Value.ToString().Trim(), 13, '0');
                    B015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B015").Value.ToString().Trim(), 13, '0');
                    B016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B016").Value.ToString().Trim(), 13, '0');
                    B017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B017").Value.ToString().Trim(), 1);
                    B018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018").Value.ToString().Trim(), 1872);

                    FileSystem.PrintLine(1, B001 + B002 + B003 + B004 + B005 + B006 + B007 + B008 + B009 + B010 + B011 + B012 + B013 + B014 + B015 + B016 + B017 + B018);

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

            // C 주(현)근무지 레코드
            string C001;    // 1     '레코드구분
            string C002;    // 2     '자료구분
            string C003;    // 3     '세무서
            string C004;    // 6     '일련번호
            string C005;    // 10    '사업자번호
            string C006;    // 2     '종(전)근무처수
            string C007;    // 1     '거주자구분코드
            string C008;    // 2     '거주지국코드
            string C009;    // 1     '외국인단일세율적용
            string C010;    // 1     '외국법인소속파견근로자여부 1,여 2,부
            string C011;    // 30    '성명
            string C012;    // 1     '내.외국인구분
            string C013;    // 13    '주민등록번호
            string C014;    // 2     '국적코드
            string C015;    // 1     '세대주여부
            string C016;    // 1     '연말정산구분
            string C017;    // 1     '사업장단위과세자여부 1여 2부   '2'
            string C018;    // 4     '종사업장일련번호   ''공란
            string C019;    // 1     '종교관련종사자여부 1.여. 2.부  '2'
            // 근무처별소득명세_주(현)근무처
            string C020;    // 10    '주현근무처-사업자번호
            string C021;    // 60    '주현근무처-근무처명
            string C022;    // 8     '근무기간 시작연월일
            string C023;    // 8     '근무기간 종료연월일
            string C024;    // 8     '감면기간 시작연월일
            string C025;    // 8     '감면기간 종료연월일
            string C026;    // 11    '급여총액
            string C027;    // 11    '상여총액
            string C028;    // 11    '인정상여
            string C029;    // 11    '주식매수선택권행사이익
            string C030;    // 11    '우리사주조합인출금
            string C031;    // 11    '임원퇴직소득금액한도초과액
            string C032;    // 11    '직무뱔명보상긐
            string C033;    // 21    '공란
            string C034;    // 11    '계
            // 주(현)근무처 비과세 및 감면소득
            string C035;    // 10    '비과세(G01:학자금)
            string C036;    // 10    '비과세(H01:무보수위원수당)
            string C037;    // 10    '비과세(H05:경호,승선수당)
            string C038;    // 10    '비과세(H06:유아,초중등)
            string C039;    // 10    '비과세(H07:고등교육법)
            string C040;    // 10    '비과세(H08:특별법)
            string C042;    // 10    '비과세(H10:기업부설연구소)
            string C041;    // 10    '비과세(H09:연구기관등)
            string C043;    // 10    '비과세(H14:보육교사근무환경개선비)
            string C044;    // 10    '비과세(H15:사립유치원수석교사.교사의인건비)
            string C045;    // 10    '비과세(H11:취재수당)
            string C046;    // 10    '비과세(H12:벽지수당)
            string C047;    // 10    '비과세(H13:재해관련급여)
            string C048;    // 10    '비과세(H16:정부공공기관지방이전기관종사자이주수당)
            string C049;    // 10    '비과세(H17:종교활동비)
            string C050;    // 10    '비과세(I01:외국정부등근로자)
            string C051;    // 10    '비과세(K01:외국주둔군인등)
            string C052;    // 10    '비과세(M01:국외근로100만원)
            string C053;    // 10    '비과세(M02:국외근로300만원)
            string C054;    // 10    '비과세(M03:국외근로)
            string C055;    // 10    '비과세(O01:야간근로수당)
            string C056;    // 10    '비과세(Q01:출산보육수당)
            string C057;    // 10    '비과세(R10:근로장학금)
            string C058;    // 10    '비과세(R11:직무발명보상금)
            string C059;    // 10    '비과세(S01:주식매수선택권)
            string C060;    // 10    '비과세(U01:벤처기업주식매수선택권)
            string C061A;   // 10    '비과세(Y02:우리사주조합인출금50%)
            string C061B;   // 10    '비과세(Y03:우리사주조합인출금75%)
            string C061C;   // 10    '비과세(Y03:우리사주조합인출금100%)
            string C062;    // 10    '비과세(Y22:전공의수련보조수당)
            string C063;    // 10    '비과세(T01:외국인기술자)
            string C064;    // 10    '비과세(T30:성과공유중소기업경영성과급)
            string C065;    // 10    '비과세(T40:중소기업핵심인력성과보상기금소득세감면)
            string C066A;   // 10    '비과세(T11:중소기업취업청년소득세감면50%)
            string C066B;   // 10    '비과세(T12:중소기업취업청년소득세감면70%)
            string C066C;   // 10    '비과세(T13:중소기업취업청년소득세감면90%)
            string C067;    // 10    '비과세(T20:조세조약상교직자감면)
            string C068;    // 10    '비과세 계
            string C069;    // 10    '감면소득 계
            // 정산명세    
            string C070;    // 11    '총급여
            string C071;    // 10    '근로소득공제
            string C072;    // 11    '근로소득금액
            // 기본공제    
            string C073;    // 8     '본인공제금액
            string C074;    // 8     '배우자공제금액
            string C075A;   // 2     '부양가족공제인원
            string C075B;   // 8     '부양가족공제금액
            // 추가공제  
            string C076A;   // 2     '경로우대공제인원
            string C076B;   // 8     '경로우대공제금액
            string C077A;   // 2     '장애자공제인원
            string C077B;   // 8     '장애자공제금액
            string C078;    // 8     '부녀자공제금액
            string C079;    // 10    '한부모공제금액
            // 연금보험료공
            string C080A;   // 10    '국민연금보험료공제_대상금액
            string C080B;   // 10    '국민연금보험료공제_공제금액
            string C081A;   // 10    '공적연금보험료공제_공무원연금_대상금액
            string C081B;   // 10    '공적연금보험료공제_공무원연금_공제금액
            string C082A;   // 10    '공적연금보험료공제_군인연금_대상금액
            string C082B;   // 10    '공적연금보험료공제_군인연금_공제금액
            string C083A;   // 10    '공적연금보험료공제_사립학교교직원연금_대상금액
            string C083B;   // 10    '공적연금보험료공제_립학교교직원연금_공제금액
            string C084A;   // 10    '공적연금보험료공제_별정우체국연금_대상금액
            string C084B;   // 10    '공적연금보험료공제_별정우체국연금_공제금액
            // 특별소득공제
            string C085A;   // 10    '보험료_건강보험료_대상금액
            string C085B;   // 10    '보험료_건강보험료_공제금액
            string C086A;   // 10    '보험료_고용보험료_대상금액
            string C086B;   // 10    '보험료_고용보험료_공제금액
            string C087A;   // 8     '주택자금_주택임차차입금 원리금상환공제금액-대출기관
            string C087B;   // 8     '주택자금_주택임차차입금 원리금상환공제금액-거주자
            string C088A;   // 8     '2011 장기주택저당차입금이자상환공제금액-15년미만
            string C088B;   // 8     '2011 장기주택저당차입금이자상환공제금액-15-29년
            string C088C;   // 8     '2011 장기주택저당차입금이자상환공제금액-30년이상
            string C089A;   // 8     '2012 이후차입분,15년이상-고정금리비거치상환대출
            string C089B;   // 8     '2012 이후차입분,15년이상-기타대출
            string C090A;   // 8     '2015 이후차입분,15년이상-고정금리이면서비거치상환대출
            string C090B;   // 8     '2015 이후차입분,15년이상-고정금리이거나비거치상환대출
            string C090C;   // 8     '2015 이후차입분,15년이상-기타대출
            string C090D;   // 8     '2015 이후차입분,10~15년-고정금리이거나비거치상환대출
            string C091;    // 11    '기부금(이월분)
            string C092;    // 11    '공란
            string C093;    // 11    '공란
            string C094;    // 11    '계  특별소득공제계
            string C095;    // 11    '차감소득금액
            // 그밖의소득공제
            string C096;    // 8     '개인연금저축소득공제
            string C097;    // 10    '소기업소상공인공제부금
            string C098;    // 10    '주택마련저축소득공제_청약저축
            string C099;    // 10    '주택마련저축소득공제_주택청약종합저축
            string C100;    // 10    '주택마련저축소득공제_근로자주택마련저축
            string C101;    // 10    '투자조합출자등소득공제
            string C102;    // 8     '신용카드등소득공제
            string C103;    // 10    '우리사주조합출연금
            string C104;    // 10    '고용유지중소기업근로자소득공제
            string C105;    // 10    '장기집합투자증권저축
            string C106;    // 10    '공란 '0'
            string C107;    // 10    '공란 '0'
            string C108;    // 11    '그밖의소득공제계
            string C109;    // 11    '소득공제종합한도초과액
            string C110;    // 11    '종합소득과세표준
            string C111;    // 10    '산출세액
            // 세액감면     
            string C112;    // 10    '소득세법
            string C113;    // 10    '조특법
            string C114;    // 10    '조특법제30조
            string C115;    // 10    '조세조약
            string C116;    // 10    '공란
            string C117;    // 10    '공란
            string C118;    // 10    '세액감면계
            // 세액공제
            string C119;    // 10    '근로소득세액공제
            string C120A;   // 2     '자녀세액공제인원
            string C120B;   // 10    '자녀세액공제
            string C121A;   // 2     '출산.입양세액공제인원
            string C121B;   // 10    '출산.입양세액공제
            string C122A;   // 10    '연금계좌_과학기술인공제_공제대상금액
            string C122B;   // 10    '연금계좌_과학기술인공제_세액공제액
            string C123A;   // 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_공제대상금액
            string C123B;   // 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_세액공제액
            string C124A;   // 10    '연금계좌_연금저축_공제대상금액
            string C124B;   // 10    '연금계좌_연금저축_세액공제액
            string C125A;   // 10    '특별세액공제_보장성보험료_공제대상금액
            string C125B;   // 10    '특별세액공제_보장성보험료_세액공제액
            string C126A;   // 10    '특별세액공제_장애인전용보험료_공제대상금액
            string C126B;   // 10    '특별세액공제_장애인전용보험료_세액공제액
            string C127A;   // 10    '특별세액공제_의료비_공제대상금액
            string C127B;   // 10    '특별세액공제_의료비_세액공제액
            string C128A;   // 10    '특별세액공제_교육비_공제대상금액
            string C128B;   // 10    '특별세액공제_교육비_세액공제액
            string C129A;   // 10    '특별세액공제_기부금_정치자금_10만원이하_공제대상금액
            string C129B;   // 10    '특별세액공제_기부금_정치자금_10만원이하_세액공제액
            string C130A;   // 11    '특별세액공제_기부금_정치자금_10만원초과_공제대상금액
            string C130B;   // 10    '특별세액공제_기부금_정치자금_10만원초과_세액공제액
            string C131A;   // 11    '특별세액공제_기부금_법정기부금_공제대상금액
            string C131B;   // 10    '특별세액공제_기부금_법정기부금_세액공제액
            string C132A;   // 11    '특별세액공제_기부금_우리사주조합기부금_공제대상금액
            string C132B;   // 10    '특별세액공제_기부금_우리사주조합기부금_세액공제액
            string C133A;   // 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체외)
            string C133B;   // 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체외)
            string C134A;   // 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체)
            string C134B;   // 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체)
            string C135;    // 11    '공란 '0'
            string C136;    // 11    '공란 '0'
            string C137;    // 10    '특별세액공제계
            string C138;    // 10    '표준세액공제
            string C139;    // 10    '납세조합공제
            string C140;    // 10    '주택차입금
            string C141;    // 10    '외국납부
            string C142A;   // 10    '월세세액공제_공제대상금액
            string C142B;   // 8     '월세세액공제_세액공제액
            string C143;    // 10    '공란 '0'
            string C144;    // 10    '공란 '0'
            string C145;    // 10    '세액공제계
            // 결정세액
            string C146A;   // 10    '소득세
            string C146B;   // 10    '지방소득세
            string C146C;   // 10    '농특세
            // 기납부세액_주(현)근무지
            string C147A;   // 10    '소득세
            string C147B;   // 10    '지방소득세
            string C147C;   // 10    '농특세
            // 납부특례세액
            string C148A;   // 10    '소득세
            string C148B;   // 10    '지방소득세
            string C148C;   // 10    '농특세
            // 차감징수세액
            string C149A_1; // 1    '소득세(기호 음수1, 양수0)
            string C149A_2; // 10   '소득세
            string C149B_1; // 1    '지방소득세(기호 음수1, 양수0)
            string C149B_2; // 10   '지방소득세
            string C149C_1; // 1    '농특세(기호 음수1, 양수0)
            string C149C_2; // 10   '농특세
            string C150;    // 248  '공란 ''

            try
            {
                // C_RECORE QUERY
                sQry = "EXEC PH_PY980_C '" + pCLTCOD + "', '" + pyyyy + "'";
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
                        C001 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C001").Value.ToString().Trim(), 1);
                        C002 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C002").Value.ToString().Trim(), 2);
                        C003 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C003").Value.ToString().Trim(), 3);
                        C004 =    codeHelpClass.GetFixedLengthStringByte(NEWCNT.ToString(), 6, '0');
                        C005 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C005").Value.ToString().Trim(), 10);
                        C006 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C006").Value.ToString().Trim(), 2, '0');
                        C007 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C007").Value.ToString().Trim(), 1);
                        C008 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C008").Value.ToString().Trim(), 2);
                        C009 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C009").Value.ToString().Trim(), 1);
                        C010 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C010").Value.ToString().Trim(), 1);
                        C011 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C011").Value.ToString().Trim(), 30);
                        C012 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C012").Value.ToString().Trim(), 1);
                        C013 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C013").Value.ToString().Trim(), 13);
                        C014 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C014").Value.ToString().Trim(), 2);
                        C015 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C015").Value.ToString().Trim(), 1);
                        C016 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C016").Value.ToString().Trim(), 1);
                        C017 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C017").Value.ToString().Trim(), 1);
                        C018 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C018").Value.ToString().Trim(), 4);
                        C019 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C019").Value.ToString().Trim(), 1);
                        C020 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C020").Value.ToString().Trim(), 10);
                        C021 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C021").Value.ToString().Trim(), 60);
                        C022 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C022").Value.ToString().Trim(), 8, '0');
                        C023 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C023").Value.ToString().Trim(), 8, '0');
                        C024 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C024").Value.ToString().Trim(), 8, '0');
                        C025 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C025").Value.ToString().Trim(), 8, '0');
                        C026 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C026").Value.ToString().Trim(), 11, '0');
                        C027 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C027").Value.ToString().Trim(), 11, '0');
                        C028 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C028").Value.ToString().Trim(), 11, '0');
                        C029 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C029").Value.ToString().Trim(), 11, '0');
                        C030 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C030").Value.ToString().Trim(), 11, '0');
                        C031 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C031").Value.ToString().Trim(), 11, '0');
                        C032 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C032").Value.ToString().Trim(), 11, '0');
                        C033 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C033").Value.ToString().Trim(), 21, '0');
                        C034 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C034").Value.ToString().Trim(), 11, '0');
                        C035 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C035").Value.ToString().Trim(), 10, '0');
                        C036 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C036").Value.ToString().Trim(), 10, '0');
                        C037 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C037").Value.ToString().Trim(), 10, '0');
                        C038 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C038").Value.ToString().Trim(), 10, '0');
                        C039 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C039").Value.ToString().Trim(), 10, '0');
                        C040 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C040").Value.ToString().Trim(), 10, '0');
                        C041 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C041").Value.ToString().Trim(), 10, '0');
                        C042 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C042").Value.ToString().Trim(), 10, '0');
                        C043 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C043").Value.ToString().Trim(), 10, '0');
                        C044 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C044").Value.ToString().Trim(), 10, '0');
                        C045 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C045").Value.ToString().Trim(), 10, '0');
                        C046 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C046").Value.ToString().Trim(), 10, '0');
                        C047 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C047").Value.ToString().Trim(), 10, '0');
                        C048 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C048").Value.ToString().Trim(), 10, '0');
                        C049 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C049").Value.ToString().Trim(), 10, '0');
                        C050 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C050").Value.ToString().Trim(), 10, '0');
                        C051 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C051").Value.ToString().Trim(), 10, '0');
                        C052 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C052").Value.ToString().Trim(), 10, '0');
                        C053 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C053").Value.ToString().Trim(), 10, '0');
                        C054 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C054").Value.ToString().Trim(), 10, '0');
                        C055 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C055").Value.ToString().Trim(), 10, '0');
                        C056 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C056").Value.ToString().Trim(), 10, '0');
                        C057 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C057").Value.ToString().Trim(), 10, '0');
                        C058 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C058").Value.ToString().Trim(), 10, '0');
                        C059 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C059").Value.ToString().Trim(), 10, '0');
                        C060 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C060").Value.ToString().Trim(), 10, '0');
                        C061A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C061A").Value.ToString().Trim(), 10, '0');
                        C061B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C061B").Value.ToString().Trim(), 10, '0');
                        C061C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C061C").Value.ToString().Trim(), 10, '0');
                        C062 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C062").Value.ToString().Trim(), 10, '0');
                        C063 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C063").Value.ToString().Trim(), 10, '0');
                        C064 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C064").Value.ToString().Trim(), 10, '0');
                        C065 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C065").Value.ToString().Trim(), 10, '0');
                        C066A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C066A").Value.ToString().Trim(), 10, '0');
                        C066B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C066B").Value.ToString().Trim(), 10, '0');
                        C066C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C066C").Value.ToString().Trim(), 10, '0');
                        C067 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C067").Value.ToString().Trim(), 10, '0');
                        C068 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C068").Value.ToString().Trim(), 10, '0');
                        C069 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C069").Value.ToString().Trim(), 10, '0');
                        C070 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C070").Value.ToString().Trim(), 11, '0');
                        C071 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C071").Value.ToString().Trim(), 10, '0');
                        C072 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C072").Value.ToString().Trim(), 11, '0');
                        C073 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C073").Value.ToString().Trim(), 8, '0');
                        C074 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C074").Value.ToString().Trim(), 8, '0');
                        C075A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C075A").Value.ToString().Trim(), 2, '0');
                        C075B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C075B").Value.ToString().Trim(), 8, '0');
                        C076A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C076A").Value.ToString().Trim(), 2, '0');
                        C076B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C076B").Value.ToString().Trim(), 8, '0');
                        C077A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C077A").Value.ToString().Trim(), 2, '0');
                        C077B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C077B").Value.ToString().Trim(), 8, '0');
                        C078 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C078").Value.ToString().Trim(), 8, '0');
                        C079 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C079").Value.ToString().Trim(), 10, '0');
                        C080A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C080A").Value.ToString().Trim(), 10, '0');
                        C080B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C080B").Value.ToString().Trim(), 10, '0');
                        C081A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C081A").Value.ToString().Trim(), 10, '0');
                        C081B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C081B").Value.ToString().Trim(), 10, '0');
                        C082A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C082A").Value.ToString().Trim(), 10, '0');
                        C082B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C082B").Value.ToString().Trim(), 10, '0');
                        C083A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C083A").Value.ToString().Trim(), 10, '0');
                        C083B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C083B").Value.ToString().Trim(), 10, '0');
                        C084A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C084A").Value.ToString().Trim(), 10, '0');
                        C084B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C084B").Value.ToString().Trim(), 10, '0');
                        C085A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C085A").Value.ToString().Trim(), 10, '0');
                        C085B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C085B").Value.ToString().Trim(), 10, '0');
                        C086A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C086A").Value.ToString().Trim(), 10, '0');
                        C086B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C086B").Value.ToString().Trim(), 10, '0');
                        C087A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C087A").Value.ToString().Trim(), 8, '0');
                        C087B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C087B").Value.ToString().Trim(), 8, '0');
                        C088A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C088A").Value.ToString().Trim(), 8, '0');
                        C088B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C088B").Value.ToString().Trim(), 8, '0');
                        C088C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C088C").Value.ToString().Trim(), 8, '0');
                        C089A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C089A").Value.ToString().Trim(), 8, '0');
                        C089B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C089B").Value.ToString().Trim(), 8, '0');
                        C090A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090A").Value.ToString().Trim(), 8, '0');
                        C090B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090B").Value.ToString().Trim(), 8, '0');
                        C090C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090C").Value.ToString().Trim(), 8, '0');
                        C090D =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090D").Value.ToString().Trim(), 8, '0');
                        C091 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C091").Value.ToString().Trim(), 11, '0');
                        C092 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C092").Value.ToString().Trim(), 11, '0');
                        C093 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C093").Value.ToString().Trim(), 11, '0');
                        C094 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C094").Value.ToString().Trim(), 11, '0');
                        C095 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C095").Value.ToString().Trim(), 11, '0');
                        C096 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C096").Value.ToString().Trim(), 8, '0');
                        C097 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C097").Value.ToString().Trim(), 10, '0');
                        C098 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C098").Value.ToString().Trim(), 10, '0');
                        C099 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C099").Value.ToString().Trim(), 10, '0');
                        C100 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C100").Value.ToString().Trim(), 10, '0');
                        C101 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C101").Value.ToString().Trim(), 10, '0');
                        C102 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C102").Value.ToString().Trim(), 8, '0');
                        C103 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C103").Value.ToString().Trim(), 10, '0');
                        C104 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C104").Value.ToString().Trim(), 10, '0');
                        C105 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C105").Value.ToString().Trim(), 10, '0');
                        C106 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C106").Value.ToString().Trim(), 10, '0');
                        C107 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C107").Value.ToString().Trim(), 10, '0');
                        C108 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C108").Value.ToString().Trim(), 11, '0');
                        C109 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C109").Value.ToString().Trim(), 11, '0');
                        C110 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C110").Value.ToString().Trim(), 11, '0');
                        C111 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C111").Value.ToString().Trim(), 10, '0');
                        C112 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C112").Value.ToString().Trim(), 10, '0');
                        C113 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C113").Value.ToString().Trim(), 10, '0');
                        C114 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C114").Value.ToString().Trim(), 10, '0');
                        C115 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C115").Value.ToString().Trim(), 10, '0');
                        C116 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C116").Value.ToString().Trim(), 10, '0');
                        C117 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C117").Value.ToString().Trim(), 10, '0');
                        C118 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C118").Value.ToString().Trim(), 10, '0');
                        C119 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C119").Value.ToString().Trim(), 10, '0');
                        C120A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C120A").Value.ToString().Trim(), 2, '0');
                        C120B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C120B").Value.ToString().Trim(), 10, '0');
                        C121A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C121A").Value.ToString().Trim(), 2, '0');
                        C121B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C121B").Value.ToString().Trim(), 10, '0');
                        C122A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C122A").Value.ToString().Trim(), 10, '0');
                        C122B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C122B").Value.ToString().Trim(), 10, '0');
                        C123A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C123A").Value.ToString().Trim(), 10, '0');
                        C123B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C123B").Value.ToString().Trim(), 10, '0');
                        C124A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C124A").Value.ToString().Trim(), 10, '0');
                        C124B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C124B").Value.ToString().Trim(), 10, '0');
                        C125A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C125A").Value.ToString().Trim(), 10, '0');
                        C125B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C125B").Value.ToString().Trim(), 10, '0');
                        C126A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C126A").Value.ToString().Trim(), 10, '0');
                        C126B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C126B").Value.ToString().Trim(), 10, '0');
                        C127A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C127A").Value.ToString().Trim(), 10, '0');
                        C127B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C127B").Value.ToString().Trim(), 10, '0');
                        C128A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C128A").Value.ToString().Trim(), 10, '0');
                        C128B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C128B").Value.ToString().Trim(), 10, '0');
                        C129A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C129A").Value.ToString().Trim(), 10, '0');
                        C129B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C129B").Value.ToString().Trim(), 10, '0');
                        C130A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C130A").Value.ToString().Trim(), 11, '0');
                        C130B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C130B").Value.ToString().Trim(), 10, '0');
                        C131A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C131A").Value.ToString().Trim(), 11, '0');
                        C131B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C131B").Value.ToString().Trim(), 10, '0');
                        C132A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C132A").Value.ToString().Trim(), 11, '0');
                        C132B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C132B").Value.ToString().Trim(), 10, '0');
                        C133A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C133A").Value.ToString().Trim(), 11, '0');
                        C133B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C133B").Value.ToString().Trim(), 10, '0');
                        C134A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C134A").Value.ToString().Trim(), 11, '0');
                        C134B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C134B").Value.ToString().Trim(), 10, '0');
                        C135 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C135").Value.ToString().Trim(), 11, '0');
                        C136 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C136").Value.ToString().Trim(), 11, '0');
                        C137 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C137").Value.ToString().Trim(), 10, '0');
                        C138 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C138").Value.ToString().Trim(), 10, '0');
                        C139 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C139").Value.ToString().Trim(), 10, '0');
                        C140 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C140").Value.ToString().Trim(), 10, '0');
                        C141 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C141").Value.ToString().Trim(), 10, '0');
                        C142A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C142A").Value.ToString().Trim(), 10, '0');
                        C142B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C142B").Value.ToString().Trim(), 8, '0');
                        C143 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C143").Value.ToString().Trim(), 10, '0');
                        C144 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C144").Value.ToString().Trim(), 10, '0');
                        C145 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C145").Value.ToString().Trim(), 10, '0');
                        C146A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C146A").Value.ToString().Trim(), 10, '0');
                        C146B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C146B").Value.ToString().Trim(), 10, '0');
                        C146C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C146C").Value.ToString().Trim(), 10, '0');
                        C147A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C147A").Value.ToString().Trim(), 10, '0');
                        C147B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C147B").Value.ToString().Trim(), 10, '0');
                        C147C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C147C").Value.ToString().Trim(), 10, '0');
                        C148A =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C148A").Value.ToString().Trim(), 10, '0');
                        C148B =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C148B").Value.ToString().Trim(), 10, '0');
                        C148C =   codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C148C").Value.ToString().Trim(), 10, '0');
                        C149A_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149A_1").Value.ToString().Trim(), 1, '0');
                        C149A_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149A_2").Value.ToString().Trim(), 10, '0');
                        C149B_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149B_1").Value.ToString().Trim(), 1, '0');
                        C149B_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149B_2").Value.ToString().Trim(), 10, '0');
                        C149C_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149C_1").Value.ToString().Trim(), 1, '0');
                        C149C_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149C_2").Value.ToString().Trim(), 10, '0');
                        C150 =    codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C150").Value.ToString().Trim(), 248);

                        FileSystem.PrintLine(1, C001 + C002 + C003 + C004 + C005 + C006 + C007 + C008 + C009 + C010 + C011 + C012 + C013 + C014 + C015 + C016 + C017 + C018 + C019 + C020 
                                              + C021 + C022 + C023 + C024 + C025 + C026 + C027 + C028 + C029 + C030 + C031 + C032 + C033 + C034 + C035 + C036 + C037 + C038 + C039 + C040
                                              + C041 + C042 + C043 + C044 + C045 + C046 + C047 + C048 + C049 + C050 + C051 + C052 + C053 + C054 + C055 + C056 + C057 + C058 + C059 + C060
                                              + C061A + C061B + C061C + C062 + C063 + C064 + C065 + C066A + C066B + C066C + C067 +  C068 + C069 + C070 + C071 + C072 + C073 + C074 + C075A + C075B + C076A + C076B + C077A + C077B + C078 + C079 + C080A + C080B
                                              + C081A + C081B + C082A + C082B + C083A + C083B + C084A + C084B + C085A + C085B + C086A + C086B + C087A + C087B + C088A + C088B + C088C + C089A + C089B + C090A + C090B + C090C + C090D + C091 + C092 + C093 + C094 + C095 + C096 + C097 + C098 + C099 + C100 
                                              + C101 + C102 + C103 + C104 + C105 + C106 + C107 + C108 + C109 + C110 + C111 + C112 + C113 + C114 + C115 + C116 + C117 + C118 + C119 + C120A + C120B
                                              + C121A + C121B + C122A + C122B + C123A + C123B + C124A + C124B + C125A + C125B + C126A + C126B + C127A + C127B + C128A + C128B + C129A + C129B + C130A + C130B + C131A + C131B + C132A + C132B + C133A + C133B + C134A + C134B + C135 + C136 + C137 + C138 + C139 + C140
                                              + C141 + C142A + C142B + C143 + C144 + C145 + C146A + C146B + C146C + C147A + C147B + C147C + C148A + C148B + C148C + C149A_1 + C149A_2 + C149B_1 + C149B_2 + C149C_1 + C149C_2 + C150 );

                        // D 레코드: 종전근무처 레코드
                        if (Conversion.Val(C006) > 0)
                        {
                            if (File_Create_D_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                            {
                                errNum = 2;
                                throw new Exception();
                            }
                        }

                        // E 레코드: 부양가족 레코드
                        if (File_Create_E_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                        {
                            errNum = 3;
                            throw new Exception();
                        }

                        // F 레코드: 연금.저축 등 소득.세액 공제명세 레코드
                        if (File_Create_F_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                        {
                            errNum = 4;
                            throw new Exception();
                        }

                        // G 레코드: 월세.거주자간 주택임차차임금 원리금 상환액 소득공제명세 레코드
                        if (File_Create_G_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                        {
                            errNum = 5;
                            throw new Exception();
                        }

                        // H 레코드: 기부조정명세 레코드
                        if (File_Create_H_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                        {
                            errNum = 6;
                            throw new Exception();
                        }

                        // I 레코드 : 해당년도 기부명세 레코드
                        if (File_Create_I_record(C_SAUP, C_YYYY, C_SABUN, C004) == false)
                        {
                            errNum = 7;
                            throw new Exception();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("D레코드(종전근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("E레코드(부양가족 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("F레코드(연금.저축 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("G레코드(월세액.주택자료 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("H레코드(기부금조정명세 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("I레코드(해당연도 기부금명세 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

            // D 종(전)근무지 레코드
            string D001; // 1    '레코드구분
            string D002; // 2    '자료구분
            string D003; // 3    '세무서
            string D004; // 6    '일련번호
            string D005; // 10   '사업자등록번호
            string D006; // 13   '소득자주민번호
            string D007; // 1    '납세조합구분
            string D008; // 60   '법인명(상호)
            string D009; // 10   '사업자등록번호
            string D010; // 8    '근무기간시작연월일
            string D011; // 8    '근무기간종료연월일
            string D012; // 8    '감면기간시작
            string D013; // 8    '감면기간종료
            string D014; // 11   '급여총액
            string D015; // 11   '상여총액
            string D016; // 11   '인정상여
            string D017; // 11   '주식매수선택권행사이익
            string D018; // 11   '우리사주조합인출금
            string D019; // 11   '임원퇴직소득금액한도초과액
            string D020; // 11   '직무발명보상금
            string D021; // 11   '공란 '0'
            string D022; // 11   '공란 '0'
            string D023; // 11   '계
            string D024; // 10   '비과세(G01:학자금)
            string D025; // 10   '비과세(H01:무보수위원수당)
            string D026; // 10   '비과세(H05:경호,승선수당)
            string D027; // 10   '비과세(H06:유아,초중등)
            string D028; // 10   '비과세(H07:고등교육법)
            string D029; // 10   '비과세(H08:특별법)
            string D030; // 10   '비과세(H09:연구기관등)
            string D031; // 10   '비과세(H10:기업부설연구소)
            string D032; // 10   '비과세(H14:보육교사근무환경개선비)
            string D033; // 10   '비과세(H15:사립유치원수석교사.교사의인건비)
            string D034; // 10   '비과세(H11:취재수당)
            string D035; // 10   '비과세(H12:벽지수당)
            string D036; // 10   '비과세(H13:재해관련급여)
            string D037; // 10   '비과세(H16:정부공공기관지방이전기관종사자이주수당)
            string D038; // 10   '비과세(H17:종교활동비)
            string D039; // 10   '비과세(I01:외국정부등근로자)
            string D040; // 10   '비과세(K01:외국주둔군인등)
            string D041; // 10   '비과세(M01:국외근로100만원)
            string D042; // 10   '비과세(M02:국외근로300만원)
            string D043; // 10   '비과세(M03:국외근로)
            string D044; // 10   '비과세(O01:야간근로수당)
            string D045; // 10   '비과세(Q01:출산보육수당)
            string D046; // 10   '비과세(R10:근로장학금)
            string D047; // 10   '비과세(R11:직무발명보상금)
            string D048; // 10   '비과세(S01:주식매수선택권)
            string D049; // 10   '비과세(U01:벤처기업주식매수선택권)
            string D050A; // 10   '비과세(Y02:우리사주조합인출금50%)
            string D050B; // 10   '비과세(Y03:우리사주조합인출금75%)
            string D050C; // 10   '비과세(Y04:우리사주조합인출금100%)
            string D051;  // 10   '비과세(Y22:전공의수련보조수당)
            string D052;  // 10   '비과세(T01:외국인기술자)
            string D053;  // 10   '비과세(T30:성과공유중소기업경영성과급)
            string D054;  // 10   '비과세(T40:중소기업핵심인력성솨보상기금수령액)
            string D055A; // 10   '비과세(T11:중소기업취업청년소득세감면50%)
            string D055B; // 10   '비과세(T12:중소기업취업청년소득세감면70%)
            string D055C; // 10   '비과세(T13:중소기업취업청년소득세감면90%)
            string D056;  // 10   '비과세(T20:조세조약상교직자감면)
            string D057;  // 10   '비과세 계
            string D058;  // 10   '감면소득 계
            string D059A; // 10   '소득세
            string D059B; // 10   '지방소득세
            string D059C; // 10   '농특세
            string D060;  // 2    '종(전)근무처일련번호 
            string D061;  // 1412 '공란

            try
            {
                // D_RECORE QUERY
                sQry = "EXEC PH_PY980_D '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
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
                        D004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0');
                        D005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D005").Value.ToString().Trim(), 10);
                        D006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D006").Value.ToString().Trim(), 13);
                        D007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D007").Value.ToString().Trim(), 1);
                        D008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D008").Value.ToString().Trim(), 60);
                        D009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D009").Value.ToString().Trim(), 10);
                        D010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D010").Value.ToString().Trim(), 8, '0');
                        D011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D011").Value.ToString().Trim(), 8, '0');
                        D012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D012").Value.ToString().Trim(), 8, '0');
                        D013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D013").Value.ToString().Trim(), 8, '0');
                        D014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D014").Value.ToString().Trim(), 11, '0');
                        D015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D015").Value.ToString().Trim(), 11, '0');
                        D016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D016").Value.ToString().Trim(), 11, '0');
                        D017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D017").Value.ToString().Trim(), 11, '0');
                        D018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D018").Value.ToString().Trim(), 11, '0');
                        D019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D019").Value.ToString().Trim(), 11, '0');
                        D020 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D020").Value.ToString().Trim(), 11, '0');
                        D021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D021").Value.ToString().Trim(), 11, '0');
                        D022 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D022").Value.ToString().Trim(), 11, '0');
                        D023 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D023").Value.ToString().Trim(), 11, '0');
                        D024 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D024").Value.ToString().Trim(), 10, '0');
                        D025 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D025").Value.ToString().Trim(), 10, '0');
                        D026 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D026").Value.ToString().Trim(), 10, '0');
                        D027 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D027").Value.ToString().Trim(), 10, '0');
                        D028 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D028").Value.ToString().Trim(), 10, '0');
                        D029 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D029").Value.ToString().Trim(), 10, '0');
                        D030 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D030").Value.ToString().Trim(), 10, '0');
                        D031 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D031").Value.ToString().Trim(), 10, '0');
                        D032 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D032").Value.ToString().Trim(), 10, '0');
                        D033 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D033").Value.ToString().Trim(), 10, '0');
                        D034 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D034").Value.ToString().Trim(), 10, '0');
                        D035 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D035").Value.ToString().Trim(), 10, '0');
                        D036 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D036").Value.ToString().Trim(), 10, '0');
                        D037 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D037").Value.ToString().Trim(), 10, '0');
                        D038 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D038").Value.ToString().Trim(), 10, '0');
                        D039 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D039").Value.ToString().Trim(), 10, '0');
                        D040 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D040").Value.ToString().Trim(), 10, '0');
                        D041 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D041").Value.ToString().Trim(), 10, '0');
                        D042 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D042").Value.ToString().Trim(), 10, '0');
                        D043 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D043").Value.ToString().Trim(), 10, '0');
                        D044 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D044").Value.ToString().Trim(), 10, '0');
                        D045 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D045").Value.ToString().Trim(), 10, '0');
                        D046 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D046").Value.ToString().Trim(), 10, '0');
                        D047 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D047").Value.ToString().Trim(), 10, '0');
                        D048 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D048").Value.ToString().Trim(), 10, '0');
                        D049 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D049").Value.ToString().Trim(), 10, '0');
                        D050A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D050A").Value.ToString().Trim(), 10, '0');
                        D050B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D050B").Value.ToString().Trim(), 10, '0');
                        D050C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D050C").Value.ToString().Trim(), 10, '0');
                        D051 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D051").Value.ToString().Trim(), 10, '0');
                        D052 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D052").Value.ToString().Trim(), 10, '0');
                        D053 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D053").Value.ToString().Trim(), 10, '0');
                        D054 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D054").Value.ToString().Trim(), 10, '0');
                        D055A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D055A").Value.ToString().Trim(), 10, '0');
                        D055B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D055B").Value.ToString().Trim(), 10, '0');
                        D055C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D055C").Value.ToString().Trim(), 10, '0');
                        D056 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D056").Value.ToString().Trim(), 10, '0');
                        D057 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D057").Value.ToString().Trim(), 10, '0');
                        D058 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D058").Value.ToString().Trim(), 10, '0');
                        D059A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D059A").Value.ToString().Trim(), 10, '0');
                        D059B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D059B").Value.ToString().Trim(), 10, '0');
                        D059C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D059C").Value.ToString().Trim(), 10, '0');
                        D060 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D060").Value.ToString().Trim(), 2);
                        D061 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D061").Value.ToString().Trim(), 1412);

                        FileSystem.PrintLine(1, D001 + D002 + D003 + D004 + D005 + D006 + D007 + D008 + D009 + D010 + D011 + D012 + D013 + D014 + D015 + D016 + D017 + D018 + D019 + D020
                                              + D021 + D022 + D023 + D024 + D025 + D026 + D027 + D028 + D029 + D030 + D031 + D032 + D033 + D034 + D035 + D036 + D037 + D038 + D039 + D040
                                              + D041 + D042 + D043 + D044 + D045 + D046 + D047 + D048 + D049 + D050A + D050B + D050C + D051 + D052 + D053 + D054 + D055A + D055B + D055C + D056 + D057 + D058 + D059A + D059B + D059C + D060
                                              + D061 );

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
        /// E 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_E_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            int i, BUYCNT, FAMCNT = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // E 소득공제명세 레코드
            string E001; // 1    '레코드구분
            string E002; // 2    '자료구분
            string E003; // 3    '세무서
            string E004; // 6    '일련번호
            string E005; // 10   '사업자등록번호
            string E006; // 13   '소득자주민등록번호

            string[] E007 = new string[5]; // 1    '관계코드
            string[] E008 = new string[5]; // 1    '내외국인구분
            string[] E009 = new string[5]; // 30   '성명
            string[] E010 = new string[5]; // 13   '주민등록번호
            string[] E011 = new string[5]; // 1    '기본공제
            string[] E012 = new string[5]; // 1    '장애자공제
            string[] E013 = new string[5]; // 1    '부녀자공제
            string[] E014 = new string[5]; // 1    '경로우대
            string[] E015 = new string[5]; // 1    '한부모공제
            string[] E016 = new string[5]; // 1    '출산.입양공제
            string[] E017 = new string[5]; // 1    '자녀공제
            string[] E018 = new string[5]; // 1    '교육비공제 1,2,3,4
            string[] E019 = new string[5]; // 10   '국세청-보험료_건강보험
            string[] E020 = new string[5]; // 10   '국세청-보험료_고용보험
            string[] E021 = new string[5]; // 10   '국세청-보험료_보장성보험
            string[] E022 = new string[5]; // 10   '국세청-보험료_장애인전용보장성보험
            string[] E023 = new string[5]; // 10   '국세청-의료비_일반
            string[] E024 = new string[5]; // 10   '국세청-의료비_난임
            string[] E025 = new string[5]; // 10   '국세청-의료비_65세이상.장애인.건강보험산정특례자
            string[] E026 = new string[5]; // 10   '국세청-의료비_실손의료보험금
            string[] E027 = new string[5]; // 10   '국세청-교육비_일반
            string[] E028 = new string[5]; // 10   '국세청-교육비_장애인특수교육비
            string[] E029 = new string[5]; // 10   '국세청-신용카드
            string[] E030 = new string[5]; // 10   '국세청-직불카드
            string[] E031 = new string[5]; // 10   '국세청-현금영수증
            string[] E032 = new string[5]; // 10   '국세청-도서.공연사용분
            string[] E033 = new string[5]; // 10   '공란
            string[] E034 = new string[5]; // 10   '국세청-전통시장사용액
            string[] E035 = new string[5]; // 10   '국세청-대중교통이용액
            string[] E036 = new string[5]; // 13   '국세청-기부금
            string[] E037 = new string[5]; // 10   '기타-보험료_건강보험
            string[] E038 = new string[5]; // 10   '기타-보험료_고용보험
            string[] E039 = new string[5]; // 10   '기타-보험료_보장성
            string[] E040 = new string[5]; // 10   '기타-보험료_장애인전용보장성
            string[] E041 = new string[5]; // 10   '기타-의료비_일반
            string[] E042 = new string[5]; // 10   '기타-의료비_난임
            string[] E043 = new string[5]; // 10   '기타-의료비_65세이상.장애인.건강보험산정특례자
            string[] E044 = new string[5]; // 10   '기타-의료비_실손의료보험금
            string[] E045 = new string[5]; // 10   '기타-교육비_일반
            string[] E046 = new string[5]; // 10   '기타-교육비_장애인특수교육비
            string[] E047 = new string[5]; // 10   '기타-신용카드
            string[] E048 = new string[5]; // 10   '기타-직불카드
            string[] E049 = new string[5]; // 10   '기타-도서.공연사용분
            string[] E050 = new string[5]; // 10   '공란
            string[] E051 = new string[5]; // 10   '기타-전통시장사용액
            string[] E052 = new string[5]; // 10   '기타-대중교통이용액
            string[] E053 = new string[5]; // 13   '기부금

            string E242;                   // 2    '부양가족레코드일련번호

            try
            {
                // E_RECORE QUERY
                sQry = "EXEC PH_PY980_E '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    BUYCNT = 0; // 가족수
                    FAMCNT = 1; // E레코드일련번호

                    // E RECORD MOVE
                    E001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E001").Value.ToString().Trim(), 1);
                    E002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E002").Value.ToString().Trim(), 2);
                    E003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E003").Value.ToString().Trim(), 3);
                    E004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0'); // C레코드의 일련번호
                    E005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E005").Value.ToString().Trim(), 10);
                    E006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E006").Value.ToString().Trim(), 13);

                    while (!oRecordSet.EoF)
                    {
                        // 초기화
                        if (BUYCNT == 0)
                        {
                            for (i = 0; i <= 4; i++)
                            {
                                E007[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E008[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E009[i] = codeHelpClass.GetFixedLengthStringByte(" ", 30);
                                E010[i] = codeHelpClass.GetFixedLengthStringByte(" ", 13);
                                E011[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E012[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E013[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E014[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E015[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E016[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E017[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                                E018[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);

                                E019[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E020[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E021[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E022[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E023[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E024[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E025[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E026[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E027[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E028[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E029[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E030[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E031[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E032[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E033[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E034[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E035[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E036[i] = codeHelpClass.GetFixedLengthStringByte("0", 13, '0');
                                E037[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E038[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E039[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E040[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E041[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E042[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E043[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E044[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E045[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E046[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E047[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E048[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E049[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E050[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E051[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E052[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E053[i] = codeHelpClass.GetFixedLengthStringByte("0", 13, '0');
                            }
                        }

                        E007[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E007").Value.ToString().Trim(), 1);
                        E008[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E008").Value.ToString().Trim(), 1);
                        E009[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E009").Value.ToString().Trim(), 30);
                        E010[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E010").Value.ToString().Trim(), 13);
                        E011[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E011").Value.ToString().Trim(), 1);
                        E012[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E012").Value.ToString().Trim(), 1);
                        E013[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E013").Value.ToString().Trim(), 1);
                        E014[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E014").Value.ToString().Trim(), 1);
                        E015[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E015").Value.ToString().Trim(), 1);
                        E016[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E016").Value.ToString().Trim(), 1);
                        E017[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E017").Value.ToString().Trim(), 1);
                        E018[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E018").Value.ToString().Trim(), 1);

                        E019[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E019").Value.ToString().Trim(), 10, '0');
                        E020[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E020").Value.ToString().Trim(), 10, '0');
                        E021[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E021").Value.ToString().Trim(), 10, '0');
                        E022[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E022").Value.ToString().Trim(), 10, '0');
                        E023[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E023").Value.ToString().Trim(), 10, '0');
                        E024[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E024").Value.ToString().Trim(), 10, '0');
                        E025[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E025").Value.ToString().Trim(), 10, '0');
                        E026[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E026").Value.ToString().Trim(), 10, '0');
                        E027[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E027").Value.ToString().Trim(), 10, '0');
                        E028[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E028").Value.ToString().Trim(), 10, '0');
                        E029[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E029").Value.ToString().Trim(), 10, '0');
                        E030[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E030").Value.ToString().Trim(), 10, '0');
                        E031[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E031").Value.ToString().Trim(), 10, '0');
                        E032[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E032").Value.ToString().Trim(), 10, '0');
                        E033[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E033").Value.ToString().Trim(), 10, '0');
                        E034[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E034").Value.ToString().Trim(), 10, '0');
                        E035[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E035").Value.ToString().Trim(), 10, '0');
                        E036[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E036").Value.ToString().Trim(), 13, '0');
                        E037[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E037").Value.ToString().Trim(), 10, '0');
                        E038[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E038").Value.ToString().Trim(), 10, '0');
                        E039[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E039").Value.ToString().Trim(), 10, '0');
                        E040[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E040").Value.ToString().Trim(), 10, '0');
                        E041[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E041").Value.ToString().Trim(), 10, '0');
                        E042[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E042").Value.ToString().Trim(), 10, '0');
                        E043[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E043").Value.ToString().Trim(), 10, '0');
                        E044[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E044").Value.ToString().Trim(), 10, '0');
                        E045[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E045").Value.ToString().Trim(), 10, '0');
                        E046[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E046").Value.ToString().Trim(), 10, '0');
                        E047[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E047").Value.ToString().Trim(), 10, '0');
                        E048[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E048").Value.ToString().Trim(), 10, '0');
                        E049[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E049").Value.ToString().Trim(), 10, '0');
                        E050[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E050").Value.ToString().Trim(), 10, '0');
                        E051[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E051").Value.ToString().Trim(), 10, '0');
                        E052[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E052").Value.ToString().Trim(), 10, '0');
                        E053[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E053").Value.ToString().Trim(), 13, '0');

                        oRecordSet.MoveNext();

                        // If BUYCNT = 4 Then    '5개면 인쇄 0 - 4
                        if (BUYCNT == 4 | oRecordSet.EoF)
                        {
                            E242 = codeHelpClass.GetFixedLengthStringByte(FAMCNT.ToString(), 2, '0'); // 일련번호
                            // E레코드 삽입
                            FileSystem.PrintLine(1, E001 + E002 + E003 + E004 + E005 + E006
                                                  + E007[0] + E008[0] + E009[0] + E010[0] + E011[0] + E012[0] + E013[0] + E014[0] + E015[0] + E016[0] + E017[0] + E018[0] + E019[0] + E020[0]
                                                  + E021[0] + E022[0] + E023[0] + E024[0] + E025[0] + E026[0] + E027[0] + E028[0] + E029[0] + E030[0] + E031[0] + E032[0] + E033[0] + E034[0]
                                                  + E035[0] + E036[0] + E037[0] + E038[0] + E039[0] + E040[0] + E041[0] + E042[0] + E043[0] + E044[0] + E045[0] + E046[0] + E047[0] + E048[0]
                                                  + E049[0] + E050[0] + E051[0] + E052[0] + E053[0]
                                                  + E007[1] + E008[1] + E009[1] + E010[1] + E011[1] + E012[1] + E013[1] + E014[1] + E015[1] + E016[1] + E017[1] + E018[1] + E019[1] + E020[1]
                                                  + E021[1] + E022[1] + E023[1] + E024[1] + E025[1] + E026[1] + E027[1] + E028[1] + E029[1] + E030[1] + E031[1] + E032[1] + E033[1] + E034[1]
                                                  + E035[1] + E036[1] + E037[1] + E038[1] + E039[1] + E040[1] + E041[1] + E042[1] + E043[1] + E044[1] + E045[1] + E046[1] + E047[1] + E048[1]
                                                  + E049[1] + E050[1] + E051[1] + E052[1] + E053[1]
                                                  + E007[2] + E008[2] + E009[2] + E010[2] + E011[2] + E012[2] + E013[2] + E014[2] + E015[2] + E016[2] + E017[2] + E018[2] + E019[2] + E020[2]
                                                  + E021[2] + E022[2] + E023[2] + E024[2] + E025[2] + E026[2] + E027[2] + E028[2] + E029[2] + E030[2] + E031[2] + E032[2] + E033[2] + E034[2]
                                                  + E035[2] + E036[2] + E037[2] + E038[2] + E039[2] + E040[2] + E041[2] + E042[2] + E043[2] + E044[2] + E045[2] + E046[2] + E047[2] + E048[2]
                                                  + E049[2] + E050[2] + E051[2] + E052[2] + E053[2]
                                                  + E007[3] + E008[3] + E009[3] + E010[3] + E011[3] + E012[3] + E013[3] + E014[3] + E015[3] + E016[3] + E017[3] + E018[3] + E019[3] + E020[3]
                                                  + E021[3] + E022[3] + E023[3] + E024[3] + E025[3] + E026[3] + E027[3] + E028[3] + E029[3] + E030[3] + E031[3] + E032[3] + E033[3] + E034[3]
                                                  + E035[3] + E036[3] + E037[3] + E038[3] + E039[3] + E040[3] + E041[3] + E042[3] + E043[3] + E044[3] + E045[3] + E046[3] + E047[3] + E048[3]
                                                  + E049[3] + E050[3] + E051[3] + E052[3] + E053[3]
                                                  + E007[4] + E008[4] + E009[4] + E010[4] + E011[4] + E012[4] + E013[4] + E014[4] + E015[4] + E016[4] + E017[4] + E018[4] + E019[4] + E020[4]
                                                  + E021[4] + E022[4] + E023[4] + E024[4] + E025[4] + E026[4] + E027[4] + E028[4] + E029[4] + E030[4] + E031[4] + E032[4] + E033[4] + E034[4]
                                                  + E035[4] + E036[4] + E037[4] + E038[4] + E039[4] + E040[4] + E041[4] + E042[4] + E043[4] + E044[4] + E045[4] + E046[4] + E047[4] + E048[4]
                                                  + E049[4] + E050[4] + E051[4] + E052[4] + E053[4] + E242);

                            // / 다음줄넘김
                            BUYCNT = 0;
                            FAMCNT = FAMCNT + 1;
                        }
                        else
                        {
                            BUYCNT = BUYCNT + 1; // 해당사원의 부양가족일련번호
                        }
                    }
                }
                else
                {
                    errNum = 1;
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("E레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
        /// F 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_F_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = true;  // 기본을 TRUE 로
            int i, SAVCNT, RCNT = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // F 연금.저축 등 소득.세액 공제명세 레코드
            string F001; // 1    '레코드구분
            string F002; // 2    '자료구분
            string F003; // 3    '세무서
            string F004; // 6    '일련번호
            string F005; // 10   '사업자등록번호
            string F006; // 13   '소득자주민등록번호

            string[] F007 = new string[15]; // 2   '소득공제구분
            string[] F008 = new string[15]; // 3   '금융기관코드
            string[] F009 = new string[15]; // 60  '금융기관상호
            string[] F010 = new string[15]; // 20  '계좌번호
            string[] F011 = new string[15]; // 10  '납입금액
            string[] F012 = new string[15]; // 10  '소득세액공제금액
            string[] F013 = new string[15]; // 4   '투자년도
            string[] F014 = new string[15]; // 1   '투자구분  조합:1, 벤처:2

            string F127; // 2    '연금.저축레코드일련번호
            string F128 = string.Empty; // 395  '공란

            try
            {
                RCNT = 1;
                // F_RECORE QUERY
                sQry = "EXEC PH_PY980_F '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    // F RECORD MOVE
                    F001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F001").Value.ToString().Trim(), 1);
                    F002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F002").Value.ToString().Trim(), 2);
                    F003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F003").Value.ToString().Trim(), 3);
                    F004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0'); // C레코드의 일련번호
                    F005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F005").Value.ToString().Trim(), 10);
                    F006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F006").Value.ToString().Trim(), 13);

                    SAVCNT = 0;
                    while (!oRecordSet.EoF)
                    {
                        // 초기화
                        if (SAVCNT == 0)
                        {
                            for (i = 0; i <= 14; i++)  // ARRY 15개 0 - 14
                            {
                                F007[i] = codeHelpClass.GetFixedLengthStringByte(" ", 2);
                                F008[i] = codeHelpClass.GetFixedLengthStringByte(" ", 3);
                                F009[i] = codeHelpClass.GetFixedLengthStringByte(" ", 60);
                                F010[i] = codeHelpClass.GetFixedLengthStringByte(" ", 20);

                                F011[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                F012[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                F013[i] = codeHelpClass.GetFixedLengthStringByte("0", 4, '0');

                                F014[i] = codeHelpClass.GetFixedLengthStringByte(" ", 1);
                            }
                        }

                        F007[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F007").Value.ToString().Trim(), 2);
                        F008[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F008").Value.ToString().Trim(), 3);
                        F009[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F009").Value.ToString().Trim(), 60);
                        F010[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F010").Value.ToString().Trim(), 20);
                        F011[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F011").Value.ToString().Trim(), 10, '0');
                        F012[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F012").Value.ToString().Trim(), 10, '0');
                        F013[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F013").Value.ToString().Trim(), 4, '0');
                        F014[SAVCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F014").Value.ToString().Trim(), 1);

                        F128 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F128").Value.ToString().Trim(), 395);

                        oRecordSet.MoveNext();

                        // If SAVCNT 가 15개나 끝이면 인쇄
                        if (SAVCNT == 14 | oRecordSet.EoF)
                        {
                            F127 = codeHelpClass.GetFixedLengthStringByte(RCNT.ToString(), 2, '0'); // 일련번호
                            //F128 = codeHelpClass.GetFixedLengthStringByte("".ToString().Trim(), 195);
                            // F 레코드 삽입
                            FileSystem.PrintLine(1, F001 + F002 + F003 + F004 + F005 + F006
                                                  + F007[0] + F008[0] + F009[0] + F010[0] + F011[0] + F012[0] + F013[0] + F014[0] + F007[1] + F008[1] + F009[1] + F010[1] + F011[1] + F012[1] + F013[1] + F014[1]
                                                  + F007[2] + F008[2] + F009[2] + F010[2] + F011[2] + F012[2] + F013[2] + F014[2] + F007[3] + F008[3] + F009[3] + F010[3] + F011[3] + F012[3] + F013[3] + F014[3]
                                                  + F007[4] + F008[4] + F009[4] + F010[4] + F011[4] + F012[4] + F013[4] + F014[4] + F007[5] + F008[5] + F009[5] + F010[5] + F011[5] + F012[5] + F013[5] + F014[5]
                                                  + F007[6] + F008[6] + F009[6] + F010[6] + F011[6] + F012[6] + F013[6] + F014[6] + F007[7] + F008[7] + F009[7] + F010[7] + F011[7] + F012[7] + F013[7] + F014[7]
                                                  + F007[8] + F008[8] + F009[8] + F010[8] + F011[8] + F012[8] + F013[8] + F014[8] + F007[9] + F008[9] + F009[9] + F010[9] + F011[9] + F012[9] + F013[9] + F014[9]
                                                  + F007[10] + F008[10] + F009[10] + F010[10] + F011[10] + F012[10] + F013[10] + F014[10] + F007[11] + F008[11] + F009[11] + F010[11] + F011[11] + F012[11] + F013[11] + F014[11]
                                                  + F007[12] + F008[12] + F009[12] + F010[12] + F011[12] + F012[12] + F013[12] + F014[12] + F007[13] + F008[13] + F009[13] + F010[13] + F011[13] + F012[13] + F013[13] + F014[13]
                                                  + F007[14] + F008[14] + F009[14] + F010[14] + F011[14] + F012[14] + F013[14] + F014[14] + F127 + F128);
                            SAVCNT = 0;
                            RCNT = RCNT + 1;
                        }
                        else
                        {
                            SAVCNT = SAVCNT + 1; // 레코드번호
                        }
                    }
                }
                else
                {
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// G 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_G_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = true;  // 기본을 TRUE 로
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // G 월세.거주자간 주택임차차임금 원리금 상환액 소득공제명세 레코드
            string G001;  // 1    '레코드구분
            string G002;  // 2    '자료구분
            string G003;  // 3    '세무서
            string G004;  // 6    '일련번호
            string G005;  // 10   '사업자번호
            string G006;  // 13   '소득자주민번호
            // 1 
            string G007;  // 60   '임대인성명(상호)1
            string G008;  // 13   '주민등록번호
            string G009;  // 1    '주택유형
            string G010;  // 5    '주택계약면적
            string G011;  // 100  '임대차계약서상주소지
            string G012;  // 8    '임대차계약기간시작
            string G013;  // 8    '임대차계약기간종료
            string G014;  // 10   '연간월세액
            string G015;  // 10   '세액공제금액
            string G016;  // 60   '대주성명
            string G017;  // 13   '대주주민등록번호
            string G018;  // 8    '금전소비대차 계약기간시작
            string G019;  // 8    '금전소비대차 계약기간종료
            string G020;  // 4    '차입금이자율
            string G021;  // 10   '원리금상환액계
            string G022;  // 10   '원금
            string G023;  // 10   '이자
            string G024;  // 10   '공제금액
            string G025;  // 60   '임대인성명(상호)
            string G026;  // 13   '주민등록번호
            string G027;  // 1    '주택유형
            string G028;  // 5    '주택계약면적
            string G029;  // 100  '임대차계약서상주소지
            string G030;  // 8    '임대차계약기간시작
            string G031;  // 8    '임대차계약기간종료
            string G032;  // 10   '전세보증금
            // 2 
            string G033;  // 60   '임대인성명2
            string G034;  // 13   '주민등록번호
            string G035;  // 1    '주택유형
            string G036;  // 5    '주택계약면적
            string G037;  // 100  '임대차계약서상주소지
            string G038;  // 8    '임대차계약기간시작
            string G039;  // 8    '임대차계약기간종료
            string G040;  // 10   '연간월세액
            string G041;  // 10   '세액공제금액
            string G042;  // 60   '대주성명
            string G043;  // 13   '대주주민등록번호
            string G044;  // 8    '금전소비대차 계약기간시작
            string G045;  // 8    '금전소비대차 계약기간종료
            string G046;  // 4    '차입금이자율
            string G047;  // 10   '원리금산환액계
            string G048;  // 10   '원금
            string G049;  // 10   '이자
            string G050;  // 10   '공제금액
            string G051;  // 60   '임대인성명
            string G052;  // 13   '주민등록번호
            string G053;  // 1    '주택유형
            string G054;  // 5    '주택계약면적
            string G055;  // 100  '임대차계약서상주소지
            string G056;  // 8    '임대차계약기간시작
            string G057;  // 8    '임대차계약기간종료
            string G058;  // 10   '전세보증금
            // 3 
            string G059;  // 60   '임대인성명3
            string G060;  // 13   '주민등록번호
            string G061;  // 1    '주택유형
            string G062;  // 5    '주택계약면적
            string G063;  // 100  '임대차계약서상주소지
            string G064;  // 8    '임대차계약기간시작
            string G065;  // 8    '임대차계약기간종료
            string G066;  // 10   '연간월세액
            string G067;  // 10   '세액공제금액
            string G068;  // 60   '대주성명
            string G069;  // 13   '대주주민등록번호
            string G070;  // 8    '금전소비대차 계약기간시작
            string G071;  // 8    '금전소비대차 계약기간종료
            string G072;  // 4    '차입금이자율
            string G073;  // 10   '원리금산환액계
            string G074;  // 10   '원금
            string G075;  // 10   '이자
            string G076;  // 10   '공제금액
            string G077;  // 60   '임대인성명
            string G078;  // 13   '주민등록번호
            string G079;  // 1    '주택유형
            string G080;  // 5    '주택계약면적
            string G081;  // 100  '임대차계약서상주소지
            string G082;  // 8    '임대차계약기간시작
            string G083;  // 8    '임대차계약기간종료
            string G084;  // 10   '전세보증금
            string G085;  // 2    '일련번호
            string G086;  // 386  '공란

            try
            {
                // G_RECORE QUERY
                sQry = "EXEC PH_PY980_G '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    // G RECORD MOVE
                    G001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G001").Value.ToString().Trim(), 1);
                    G002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G002").Value.ToString().Trim(), 2);
                    G003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G003").Value.ToString().Trim(), 3);
                    G004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0'); // C레코드의 일련번호
                    G005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G005").Value.ToString().Trim(), 10);
                    G006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G006").Value.ToString().Trim(), 13);

                    G007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G007").Value.ToString().Trim(), 60);
                    G008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G008").Value.ToString().Trim(), 13);
                    G009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G009").Value.ToString().Trim(), 1);
                    G010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G010").Value.ToString().Trim(), 5, '0');
                    G011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G011").Value.ToString().Trim(), 100);
                    G012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G012").Value.ToString().Trim(), 8, '0');
                    G013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G013").Value.ToString().Trim(), 8, '0');
                    G014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G014").Value.ToString().Trim(), 10, '0');
                    G015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G015").Value.ToString().Trim(), 10, '0');
                    G016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G016").Value.ToString().Trim(), 60);
                    G017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G017").Value.ToString().Trim(), 13);
                    G018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G018").Value.ToString().Trim(), 8, '0');
                    G019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G019").Value.ToString().Trim(), 8, '0'); ;
                    G020 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G020").Value.ToString().Trim(), 4, '0');
                    G021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G021").Value.ToString().Trim(), 10, '0');
                    G022 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G022").Value.ToString().Trim(), 10, '0');
                    G023 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G023").Value.ToString().Trim(), 10, '0');
                    G024 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G024").Value.ToString().Trim(), 10, '0');
                    G025 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G025").Value.ToString().Trim(), 60);
                    G026 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G026").Value.ToString().Trim(), 13);
                    G027 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G027").Value.ToString().Trim(), 1);
                    G028 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G028").Value.ToString().Trim(), 5, '0');
                    G029 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G029").Value.ToString().Trim(), 100);
                    G030 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G030").Value.ToString().Trim(), 8, '0');
                    G031 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G031").Value.ToString().Trim(), 8, '0');
                    G032 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G032").Value.ToString().Trim(), 10, '0');

                    G033 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G033").Value.ToString().Trim(), 60);
                    G034 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G034").Value.ToString().Trim(), 13);
                    G035 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G035").Value.ToString().Trim(), 1);
                    G036 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G036").Value.ToString().Trim(), 5, '0');
                    G037 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G037").Value.ToString().Trim(), 100);
                    G038 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G038").Value.ToString().Trim(), 8, '0');
                    G039 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G039").Value.ToString().Trim(), 8, '0');
                    G040 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G040").Value.ToString().Trim(), 10, '0');
                    G041 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G041").Value.ToString().Trim(), 10, '0');
                    G042 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G042").Value.ToString().Trim(), 60);
                    G043 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G043").Value.ToString().Trim(), 13);
                    G044 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G044").Value.ToString().Trim(), 8, '0');
                    G045 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G045").Value.ToString().Trim(), 8, '0');
                    G046 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G046").Value.ToString().Trim(), 4, '0');
                    G047 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G047").Value.ToString().Trim(), 10, '0');
                    G048 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G048").Value.ToString().Trim(), 10, '0');
                    G049 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G049").Value.ToString().Trim(), 10, '0');
                    G050 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G050").Value.ToString().Trim(), 10, '0');
                    G051 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G051").Value.ToString().Trim(), 60);
                    G052 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G052").Value.ToString().Trim(), 13);
                    G053 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G053").Value.ToString().Trim(), 1);
                    G054 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G054").Value.ToString().Trim(), 5, '0');
                    G055 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G055").Value.ToString().Trim(), 100);
                    G056 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G056").Value.ToString().Trim(), 8, '0');
                    G057 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G057").Value.ToString().Trim(), 8, '0');
                    G058 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G058").Value.ToString().Trim(), 10, '0');

                    G059 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G059").Value.ToString().Trim(), 60);
                    G060 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G060").Value.ToString().Trim(), 13);
                    G061 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G061").Value.ToString().Trim(), 1);
                    G062 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G062").Value.ToString().Trim(), 5, '0');
                    G063 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G063").Value.ToString().Trim(), 100);
                    G064 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G064").Value.ToString().Trim(), 8, '0');
                    G065 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G065").Value.ToString().Trim(), 8, '0');
                    G066 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G066").Value.ToString().Trim(), 10, '0');
                    G067 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G067").Value.ToString().Trim(), 10, '0');
                    G068 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G068").Value.ToString().Trim(), 60);
                    G069 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G069").Value.ToString().Trim(), 13);
                    G070 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G070").Value.ToString().Trim(), 8, '0');
                    G071 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G071").Value.ToString().Trim(), 8, '0');
                    G072 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G072").Value.ToString().Trim(), 4, '0');
                    G073 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G073").Value.ToString().Trim(), 10, '0');
                    G074 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G074").Value.ToString().Trim(), 10, '0');
                    G075 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G075").Value.ToString().Trim(), 10, '0');
                    G076 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G076").Value.ToString().Trim(), 10, '0');
                    G077 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G077").Value.ToString().Trim(), 60);
                    G078 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G078").Value.ToString().Trim(), 13);
                    G079 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G079").Value.ToString().Trim(), 1);
                    G080 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G080").Value.ToString().Trim(), 5, '0');
                    G081 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G081").Value.ToString().Trim(), 100);
                    G082 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G082").Value.ToString().Trim(), 8, '0');
                    G083 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G083").Value.ToString().Trim(), 8, '0');
                    G084 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G084").Value.ToString().Trim(), 10, '0');
                    G085 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G085").Value.ToString().Trim(), 2);
                    G086 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G086").Value.ToString().Trim(), 386);

                    // G 레코드 삽입
                    FileSystem.PrintLine(1, G001 + G002 + G003 + G004 + G005 + G006 + G007 + G008 + G009 + G010 + G011 + G012 + G013 + G014 + G015 + G016 + G017 + G018 + G019 + G020
                                          + G021 + G022 + G023 + G024 + G025 + G026 + G027 + G028 + G029 + G030 + G031 + G032 + G033 + G034 + G035 + G036 + G037 + G038 + G039 + G040
                                          + G041 + G042 + G043 + G044 + G045 + G046 + G047 + G048 + G049 + G050 + G051 + G052 + G053 + G054 + G055 + G056 + G057 + G058 + G059 + G060
                                          + G061 + G062 + G063 + G064 + G065 + G066 + G067 + G068 + G069 + G070 + G071 + G072 + G073 + G074 + G075 + G076 + G077 + G078 + G079 + G080
                                          + G081 + G082 + G083 + G084 + G085 + G086);
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// H 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_H_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = true;  // 기본을 TRUE 로
            int HCount = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // H 기부조정명세 레코드
            string H001;  // 1     '레코드구분
            string H002;  // 2     '자료구분
            string H003;  // 3     '세무서
            string H004;  // 6     '일련번호
            string H005;  // 10    '사업자번호
            string H006;  // 13    '소득자주민등록번호
            string H007;  // 1     '내,외국인코드
            string H008;  // 30    '성명
            string H009;  // 2     '유형코드
            string H010;  // 4     '기부년도
            string H011;  // 13    '기부금액
            string H012;  // 13    '전년까지공제된금액
            string H013;  // 13    '공제대상금액
            string H014;  // 13    '해당년도공제금액 필요경비 '0'  2016
            string H015;  // 13    '해당년도공제금액세액(소득)공제금액  2016
            string H016;  // 13    '해당년도에공제받지못한금액_소멸금액
            string H017;  // 13    '해당년도에공제받지못한금액_이월금액
            string H018;  // 5     '기부조정명세일련번호
            string H019;  // 1914  '공란

            try
            {
                // H_RECORE QUERY
                sQry = "EXEC PH_PY980_H '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    HCount = 0;
                    while (!oRecordSet.EoF)
                    {
                        HCount = HCount + 1; // 일련번호
                        // H RECORD MOVE
                        H001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H001").Value.ToString().Trim(), 1);
                        H002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H002").Value.ToString().Trim(), 2);
                        H003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H003").Value.ToString().Trim(), 3);
                        H004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0'); // C레코드의 일련번호
                        H005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H005").Value.ToString().Trim(), 10);
                        H006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H006").Value.ToString().Trim(), 13);
                        H007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H007").Value.ToString().Trim(), 1);
                        H008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H008").Value.ToString().Trim(), 30);
                        H009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H009").Value.ToString().Trim(), 2);
                        H010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H010").Value.ToString().Trim(), 4);
                        H011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H011").Value.ToString().Trim(), 13, '0');
                        H012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H012").Value.ToString().Trim(), 13, '0');
                        H013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H013").Value.ToString().Trim(), 13, '0');
                        H014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H014").Value.ToString().Trim(), 13, '0');
                        H015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H015").Value.ToString().Trim(), 13, '0');
                        H016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H016").Value.ToString().Trim(), 13, '0');
                        H017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H017").Value.ToString().Trim(), 13, '0');
                        H018 = codeHelpClass.GetFixedLengthStringByte(HCount.ToString().Trim(), 5, '0');
                        H019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H019").Value.ToString().Trim(), 1914);

                        // H 레코드 삽입
                        FileSystem.PrintLine(1, H001 + H002 + H003 + H004 + H005 + H006 + H007 + H008 + H009 + H010
                                              + H011 + H012 + H013 + H014 + H015 + H016 + H017 + H018 + H019);
                       oRecordSet.MoveNext();
                    }
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// I 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_I_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool functionReturnValue = true;  // 기본을 TRUE 로
            int ICount = 0;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // I 해당년도 기부명세 레코드
            string I001;  // 1    '레코드구분
            string I002;  // 2    '자료구분
            string I003;  // 3    '세무서
            string I004;  // 6    '일련번호
            string I005;  // 10   '사업자등록번호
            string I006;  // 13   '주민등록번호
            string I007;  // 2    '유형코드
            string I008;  // 1    '기부내용
            string I009;  // 13   '기부처-사업자등록번호
            string I010;  // 60   '기부처-법인명(상호)
            string I011;  // 1    '관계
            string I012;  // 1    '내,외국인코드
            string I013;  // 30   '성명
            string I014;  // 13   '주민등록번호
            string I015;  // 5    '건수
            string I016;  // 13   '기부금합계금액
            string I017;  // 13   '공제대상기부금액
            string I018;  // 13   '기부장려금신청금액
            string I019;  // 13   '기타
            string I020;  // 5    '해당연도기부명세일련번호
            string I021;  // 1864 '공란

            try
            {
                // H_RECORE QUERY
                sQry = "EXEC PH_PY980_I '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    ICount = 0;
                    while (!oRecordSet.EoF)
                    {
                        ICount = ICount + 1; // 일련번호
                        // I RECORD MOVE
                        I001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I001").Value.ToString().Trim(), 1);
                        I002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I002").Value.ToString().Trim(), 2);
                        I003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I003").Value.ToString().Trim(), 3);
                        I004 = codeHelpClass.GetFixedLengthStringByte(pC004.ToString(), 6, '0'); // C레코드의 일련번호
                        I005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I005").Value.ToString().Trim(), 10);
                        I006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I006").Value.ToString().Trim(), 13);
                        I007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I007").Value.ToString().Trim(), 2);
                        I008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I008").Value.ToString().Trim(), 1);
                        I009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I009").Value.ToString().Trim(), 13);
                        I010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I010").Value.ToString().Trim(), 60);
                        I011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I011").Value.ToString().Trim(), 1);
                        I012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I012").Value.ToString().Trim(), 1);
                        I013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I013").Value.ToString().Trim(), 30);
                        I014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I014").Value.ToString().Trim(), 13);
                        I015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I015").Value.ToString().Trim(), 5, '0');
                        I016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I016").Value.ToString().Trim(), 13, '0');
                        I017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I017").Value.ToString().Trim(), 13, '0');
                        I018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I018").Value.ToString().Trim(), 13, '0');
                        I019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I019").Value.ToString().Trim(), 13, '0');
                        I020 = codeHelpClass.GetFixedLengthStringByte(ICount.ToString().Trim(), 5, '0');
                        I021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I021").Value.ToString().Trim(), 1864);

                        // I 레코드 삽입
                        FileSystem.PrintLine(1, I001 + I002 + I003 + I004 + I005 + I006 + I007 + I008 + I009 + I010
                                              + I011 + I012 + I013 + I014 + I015 + I016 + I017 + I018 + I019 + I020 + I021 );
                        oRecordSet.MoveNext();
                    }
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

//internal class PH_PY980
//{
//    // //  SAP MANAGE UI API 2004 SDK Sample
//    // //****************************************************************************
//    // //  File           : PH_PY980.cls
//    // //  Module         : 인사관리>정산관리
//    // //  Desc           : 근로소득지급명세서자료 전산매체수록
//    // //  FormType       :
//    // //  Create Date    : 2014.01.17
//    // //  Modified Date  : 2015.01.21
//    // //  Creator        : NGY
//    // //  Modifier       :
//    // //  Copyright  (c) Poongsan Holdings
//    // //****************************************************************************


//    public string oFormUniqueID;
//    public SAPbouiCOM.Form oForm;
//    private SAPbobsCOM.Recordset sRecordset;
//    private SAPbouiCOM.Matrix oMat1;
//    private string Last_Item; // 클래스에서 선택한 마지막 아이템 Uid값

//    private string CLTCOD;
//    private string yyyy;
//    private string HtaxID;
//    private string TeamName;
//    private string Dname;
//    private string Dtel;
//    private string DocDate;
//    private string oFilePath;

//    private VB6.FixedLengthString FILNAM = new VB6.FixedLengthString(30); // 파  일  명
//    private int MaxRow;
//    private short BUSCNT; // / B레코드일련번호
//    private short BUSTOT; // / B레코드총갯수

//    private short NEWCNT;
//    private short OLDCNT;
//    private string C_SAUP;
//    private string C_YYYY;
//    private string C_SABUN;
//    private string E_BUYCNT;
//    private string C_BUYCNT;

//    // 2013년기준 1400 BYTE
//    // 2014년기준 1520 BYTE
//    // 2014년기준 1580 BYTE  re
//    // 2015년귀속 1610 BYTE
//    // 2016년귀속 1620 BYTE
//    // 2017년귀속 1620 BYTE
//    // 2018년귀속 1882 BYTE

//    // / A 제출자 레코드
//    private struct A_record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] A001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] A002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] A003; // 세무서코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] A004; // 제출일자
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] A005; // 제출자구분 (1;세무대리인, 2;법인, 3;개인)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] A006; // 세무대리인
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] A007; // 홈텍스ID
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] A008; // 세무프로그램코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] A009; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] A010; // 법인명(상호)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] A011; // 담당자부서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] A012; // 담당자성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(15)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 15)]
//        public char[] A013; // 담당자전화번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] A014; // 신고의무자수
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] A015; // 한글코드종류
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1684)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1684)]
//        public char[] A016; // 공란
//    }
//    private A_record A_rec;

//    // / B 원천징수의무자별 집계 레코드
//    private struct B_record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] B001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] B002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] B003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] B004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] B005; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] B006; // 법인명(상호)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] B007; // 대표자
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] B008; // 주민(법인)번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(7)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//        public char[] B009; // 주(현)근무처(C레코드)수
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(7)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 7)]
//        public char[] B010; // 종(전)근무처(D레코드)수
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(14)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
//        public char[] B011; // 총급여총계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] B012; // 결정세액(소득세)총계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] B013; // 결정세액(지방소득세)총계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] B014; // 결정세액(농특세)총계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] B015; // 결정세액총계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] B016; // 제출대상기간
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1676)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1676)]
//        public char[] B017; // 공란
//    }
//    private B_record B_rec;

//    // / C 주(현)근무지 레코드
//    private struct C_record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] C003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] C004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C005; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C006; // 종(전)근무처수
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C007; // 거주자구분코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C008; // 거주지국코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C009; // 외국인단일세율적용
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C010; // 외국법인소속파견근로자여부 1,여 2,부
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] C011; // 성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C012; // 내.외국인구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] C013; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C014; // 국적코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C015; // 세대주여부
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C016; // 연말정산구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C017; // 사업장단위과세자여부 1여 2부   '2'
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] C018; // 종사업장일련번호   ''공란
//                            // /근무처별소득명세_주(현)근무처
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C019; // 주현근무처-사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] C020; // 주현근무처-근무처명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C021; // 근무기간 시작연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C022; // 근무기간 종료연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C023; // 감면기간 시작연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C024; // 감면기간 종료연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C025; // 급여총액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C026; // 상여총액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C027; // 인정상여
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C028; // 주식매수선택권행사이익
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C029; // 우리사주조합인출금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C030; // 임원퇴직소득금액한도초과액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C031; // 직무u명보상O
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] C032; // 공란
//                            // 2018년 C033 없슴
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C034; // 계
//                            // /주(현)근무처 비과세 및 감면소득
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C035; // 비과세(G01:학자금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C036; // 비과세(H01:무보수위원수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C037; // 비과세(H05:경호,승선수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C038; // 비과세(H06:유아,초중등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C039; // 비과세(H07:고등교육법)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C040; // 비과세(H08:특별법)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C041; // 비과세(H09:연구기관등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C042; // 비과세(H10:기업부설연구소)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C043; // 비과세(H14:보육교사근무환경개선비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C044; // 비과세(H15:사립유치원수석교사.교사의인건비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C045; // 비과세(H11:취재수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C046; // 비과세(H12:벽지수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C047; // 비과세(H13:재해관련급여)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C048; // 비과세(H16:정부공공기관지방이전기관종사자이주수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C049; // 비과세(H17:종교활동비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C050; // 비과세(I01:외국정부등근로자)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C051; // 비과세(K01:외국주둔군인등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C052; // 비과세(M01:국외근로100만원)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C053; // 비과세(M02:국외근로300만원)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C054; // 비과세(M03:국외근로)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C055; // 비과세(O01:야간근로수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C056; // 비과세(Q01:출산보육수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C057; // 비과세(R10:근로장학금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C058; // 비과세(R11:직무발명보상금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C059; // 비과세(S01:주식매수선택권)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C060; // 비과세(U01:벤처기업주식매수선택권)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C061A; // 비과세(Y02:우리사주조합인출금50%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C061B; // 비과세(Y03:우리사주조합인출금75%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C061C; // 비과세(Y03:우리사주조합인출금100%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C062; // 비과세(Y22:전공의수련보조수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C063; // 비과세(T01:외국인기술자)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C064A; // 비과세(T10:중소기업취업청년소득세감면100%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C064B; // 비과세(T11:중소기업취업청년소득세감면50%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C064C; // 비과세(T12:중소기업취업청년소득세감면70%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C064D; // 비과세(T13:중소기업취업청년소득세감면90%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C065; // 비과세(T20:조세조약상교직자감면)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] C066; // 공란  '0'
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C068; // 비과세 계
//                            // 2018년 C067 없슴
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C069; // 감면소득 계
//                            // /정산명세
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C070; // 총급여
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C071; // 근로소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C072; // 근로소득금액
//                            // /기본공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C073; // 본인공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C074; // 배우자공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C075A; // 부양가족공제인원
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C075B; // 부양가족공제금액
//                             // /추가공제
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C076A; // 경로우대공제인원
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C076B; // 경로우대공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C077A; // 장애자공제인원
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C077B; // 장애자공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C078; // 부녀자공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C079; // 한부모공제금액
//                            // /연금보험료공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C080A; // 국민연금보험료공제_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C080B; // 국민연금보험료공제_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C081A; // 공적연금보험료공제_공무원연금_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C081B; // 공적연금보험료공제_공무원연금_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C082A; // 공적연금보험료공제_군인연금_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C082B; // 공적연금보험료공제_군인연금_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C083A; // 공적연금보험료공제_사립학교교직원연금_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C083B; // 공적연금보험료공제_립학교교직원연금_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C084A; // 공적연금보험료공제_별정우체국연금_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C084B; // 공적연금보험료공제_별정우체국연금_공제금액
//                             // /특별소득공제
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C085A; // 보험료_건강보험료_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C085B; // 보험료_건강보험료_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C086A; // 보험료_고용보험료_대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C086B; // 보험료_고용보험료_공제금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C087A; // 주택자금_주택임차차입금 원리금상환공제금액-대출기관
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C087B; // 주택자금_주택임차차입금 원리금상환공제금액-거주자
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C088A; // 2011 장기주택저당차입금이자상환공제금액-15년미만
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C088B; // 2011 장기주택저당차입금이자상환공제금액-15-29년
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C088C; // 2011 장기주택저당차입금이자상환공제금액-30년이상
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C089A; // 2012 이후차입분,15년이상-고정금리비거치상환대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C089B; // 2012 이후차입분,15년이상-기타대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C090A; // 2015 이후차입분,15년이상-고정금리이면서비거치상환대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C090B; // 2015 이후차입분,15년이상-고정금리이거나비거치상환대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C090C; // 2015 이후차입분,15년이상-기타대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C090D; // 2015 이후차입분,10~15년-고정금리이거나비거치상환대출
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C091; // 기부금(이월분)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] C092; // 공란
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C094; // 계  특별소득공제계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C095; // 차감소득금액
//                            // /그밖의소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C096; // 개인연금저축소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C097; // 소기업소상공인공제부금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C098; // 주택마련저축소득공제_청약저축
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C099; // 주택마련저축소득공제_주택청약종합저축
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C100; // 주택마련저축소득공제_근로자주택마련저축
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C101; // 투자조합출자등소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C102; // 신용카드등소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C103; // 우리사주조합출연금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C104; // 고용유지중소기업근로자소득공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C105; // 장기집합투자증권저축
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] C106; // 공란 '0'
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C108; // 그밖의소득공제계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C109; // 소득공제종합한도초과액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C110; // 종합소득과세표준
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C111; // 산출세액
//                            // /세액감면
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C112; // 소득세법
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C113; // 조특법
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C114; // 조특법제30조
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C115; // 조세조약
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] C116; // 공란
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C118; // 세액감면계
//                            // /세액공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C119; // 근로소득세액공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C120A; // 자녀세액공제인원
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C120B; // 자녀세액공제
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] C121A; // 출산.입양세액공제인원
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C121B; // 출산.입양세액공제
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C122A; // 연금계좌_과학기술인공제_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C122B; // 연금계좌_과학기술인공제_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C123A; // 연금계좌_근로자퇴직급여보장법에따른 퇴직급여_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C123B; // 연금계좌_근로자퇴직급여보장법에따른 퇴직급여_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C124A; // 연금계좌_연금저축_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C124B; // 연금계좌_연금저축_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C125A; // 특별세액공제_보장성보험료_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C125B; // 특별세액공제_보장성보험료_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C126A; // 특별세액공제_장애인전용보험료_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C126B; // 특별세액공제_장애인전용보험료_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C127A; // 특별세액공제_의료비_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C127B; // 특별세액공제_의료비_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C128A; // 특별세액공제_교육비_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C128B; // 특별세액공제_교육비_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C129A; // 특별세액공제_기부금_정치자금_10만원이하_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C129B; // 특별세액공제_기부금_정치자금_10만원이하_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C130A; // 특별세액공제_기부금_정치자금_10만원초과_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C130B; // 특별세액공제_기부금_정치자금_10만원초과_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C131A; // 특별세액공제_기부금_법정기부금_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C131B; // 특별세액공제_기부금_법정기부금_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C132A; // 특별세액공제_기부금_우리사주조합기부금_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C132B; // 특별세액공제_기부금_우리사주조합기부금_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C133A; // 특별세액공제_기부금_지정기부금_공제대상금액(종교단체외)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C133B; // 특별세액공제_기부금_지정기부금_세액공제액(종교단체외)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] C134A; // 특별세액공제_기부금_지정기부금_공제대상금액(종교단체)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C134B; // 특별세액공제_기부금_지정기부금_세액공제액(종교단체)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] C135; // 공란 '0'
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C137; // 특별세액공제계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C138; // 표준세액공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C139; // 납세조합공제
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C140; // 주택차입금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C141; // 외국납부
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C142A; // 월세세액공제_공제대상금액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] C142B; // 월세세액공제_세액공제액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] C143; // 공란 '0'
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C145; // 세액공제계
//                            // /결정세액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C146A; // 소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C146B; // 지방소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C146C; // 농특세
//                             // /기납부세액_주(현)근무지
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C147A; // 소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C147B; // 지방소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C147C; // 농특세
//                             // /납부특례세액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C148A; // 소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C148B; // 지방소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C148C; // 농특세
//                             // /차감징수세액
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C149A_1; // 소득세(기호 음수1, 양수0)
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C149A_2; // 소득세
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C149B_1; // 지방소득세(기호 음수1, 양수0)
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C149B_2; // 지방소득세
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] C149C_1; // 농특세(기호 음수1, 양수0)
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] C149C_2; // 농특세
//                               // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(38)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 38)]
//        public char[] C150; // 공란 ''
//    }
//    private C_record C_rec;

//    // / D 종(전)근무지 레코드
//    private struct D_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] D001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] D002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] D003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] D004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D005; // 사업자등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] D006; // 소득자주민번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] D007; // 납세조합구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] D008; // 법인명(상호)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D009; // 사업자등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] D010; // 근무기간시작연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] D011; // 근무기간종료연월일
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] D012; // 감면기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] D013; // 감면기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D014; // 급여총액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D015; // 상여총액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D016; // 인정상여
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D017; // 주식매수선택권행사이익
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D018; // 우리사주조합인출금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D019; // 임원퇴직소득금액한도초과액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D020; // 직무발명보상금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(22)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 22)]
//        public char[] D021; // 공란
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(11)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 11)]
//        public char[] D023; // 계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D024; // 비과세(G01:학자금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D025; // 비과세(H01:무보수위원수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D026; // 비과세(H05:경호,승선수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D027; // 비과세(H06:유아,초중등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D028; // 비과세(H07:고등교육법)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D029; // 비과세(H08:특별법)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D030; // 비과세(H09:연구기관등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D031; // 비과세(H10:기업부설연구소)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D032; // 비과세(H14:보육교사근무환경개선비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D033; // 비과세(H15:사립유치원수석교사.교사의인건비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D034; // 비과세(H11:취재수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D035; // 비과세(H12:벽지수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D036; // 비과세(H13:재해관련급여)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D037; // 비과세(H16:정부공공기관지방이전기관종사자이주수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D038; // 비과세(H17:종교활동비)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D039; // 비과세(I01:외국정부등근로자)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D040; // 비과세(K01:외국주둔군인등)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D041; // 비과세(M01:국외근로100만원)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D042; // 비과세(M02:국외근로300만원)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D043; // 비과세(M03:국외근로)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D044; // 비과세(O01:야간근로수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D045; // 비과세(Q01:출산보육수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D046; // 비과세(R10:근로장학금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D047; // 비과세(R11:직무발명보상금)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D048; // 비과세(S01:주식매수선택권)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D049; // 비과세(U01:벤처기업주식매수선택권)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D050A; // 비과세(Y02:우리사주조합인출금50%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D050B; // 비과세(Y03:우리사주조합인출금75%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D050C; // 비과세(Y04:우리사주조합인출금100%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D051; // 비과세(Y22:전공의수련보조수당)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D052; // 비과세(T01:외국인기술자)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D053A; // 비과세(T10:중소기업취업청년소득세감면100%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D053B; // 비과세(T11:중소기업취업청년소득세감면50%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D053C; // 비과세(T12:중소기업취업청년소득세감면70%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D053D; // 비과세(T13:중소기업취업청년소득세감면90%)
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D054; // 비과세(T20:조세조약상교직자감면)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(20)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//        public char[] D055; // 공란 '0'기재
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D057; // 비과세 계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D058; // 감면소득 계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D059A; // 소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D059B; // 지방소득세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] D059C; // 농특세
//                             // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] D060; // 종(전)근무처일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1202)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1202)]
//        public char[] D061; // 공란
//    }
//    private D_Record D_rec;

//    // / E 소득공제명세 레코드
//    private struct E_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] E001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] E002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] E003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] E004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] E005; // 사업자등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] E006; // 소득자주민등록번호
//                            // ARRY 5   E007 - E146
//                            // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E007 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 관계코드
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E008 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 내외국인구분
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E009 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 성명
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E010 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 주민등록번호
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E011 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기본공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E012 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 장애자공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E013 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 부녀자공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E014 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 경로우대
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E015 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 한부모공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E016 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 출산.입양공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E017 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 자녀공제
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E018 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 교육비공제 1,2,3,4
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E019 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-보험료_건강보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E020 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-보험료_고용보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E021 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-보험료_보장성보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E022 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-보험료_장애인전용보장성보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E023 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-의료비_일반
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E024 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-의료비_난임
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E025 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-의료비_장애인.건강보험산정특례자
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E026 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-교육비_일반
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E027 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-교육비_장애인특수교육비
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E028 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-신용카드
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E029 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-직불카드
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E030 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-현금영수증
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E031 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-도서.공연사용분
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E032 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-전통시장사용액
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E033 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-대중교통이용액
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E034 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 국세청-기부금
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E035 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-보험료_건강보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E036 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-보험료_고용보험
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E037 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-보험료_보장성
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E038 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-보험료_장애인전용보장성
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E039 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-의료비_일반
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E040 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-의료비_난임
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E041 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-의료비_장애인.건강보험산정특례자
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E042 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-교육비_일반
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E043 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-교육비_장애인특수교육비
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E044 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-신용카드
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E045 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-직불카드
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E046 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-도서.공연사용분
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E047 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-전통시장사용액
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E048 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기타-대중교통이용액
//                                                                                                                                        // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] E049 = new string[6];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 기부금
//                                                                                                                                        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] E222; // 부양가족레코드일련번호
//    }
//    // UPGRADE_WARNING: E_rec 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//    private E_Record E_rec;

//    // / F 연금.저축 등 소득.세액 공제명세 레코드
//    private struct F_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] F001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] F002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] F003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] F004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] F005; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] F006; // 소득자주민번호
//                            // ARRY 15   F007 - F096
//                            // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F007 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 소득공제구분
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F008 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 금융기관코드
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F009 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 금융기관상호
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F010 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 계좌번호
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F011 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 납입금액
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F012 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 소득세액공제금액
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F013 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 투자년도
//                                                                                                                                         // UPGRADE_ISSUE: 지원되지 않는 선언 형식입니다. 고정 길이 문자열 배열입니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
//        public string[] F014 = new string[16];/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */ // 투자구분  조합:1, 벤처:2
//                                                                                                                                         // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] F127; // 연금.저축레코드일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(195)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 195)]
//        public char[] F128; // 공란
//    }
//    // UPGRADE_WARNING: F_rec 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//    private F_Record F_rec;

//    // / G 월세.거주자간 주택임차차임금 원리금 상환액 소득공제명세 레코드
//    private struct G_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] G002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] G003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] G004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G005; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G006; // 소득자주민번호
//                            // /1
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G007; // 임대인성명(상호)1
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G008; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G009; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G010; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G011; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G012; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G013; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G014; // 연간월세액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G015; // 세액공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G016; // 대주성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G017; // 대주주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G018; // 금전소비대차 계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G019; // 금전소비대차 계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] G020; // 차입금이자율
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G021; // 원리금상환액계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G022; // 원금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G023; // 이자
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G024; // 공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G025; // 임대인성명(상호)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G026; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G027; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G028; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G029; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G030; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G031; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G032; // 전세보증금
//                            // /2
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G033; // 임대인성명2
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G034; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G035; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G036; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G037; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G038; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G039; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G040; // 연간월세액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G041; // 세액공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G042; // 대주성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G043; // 대주주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G044; // 금전소비대차 계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G045; // 금전소비대차 계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] G046; // 차입금이자율
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G047; // 원리금산환액계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G048; // 원금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G049; // 이자
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G050; // 공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G051; // 임대인성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G052; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G053; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G054; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G055; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G056; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G057; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G058; // 전세보증금
//                            // /3
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G059; // 임대인성명3
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G060; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G061; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G062; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G063; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G064; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G065; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G066; // 연간월세액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G067; // 세액공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G068; // 대주성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G069; // 대주주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G070; // 금전소비대차 계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G071; // 금전소비대차 계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] G072; // 차입금이자율
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G073; // 원리금산환액계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G074; // 원금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G075; // 이자
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G076; // 공제금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] G077; // 임대인성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] G078; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] G079; // 주택유형
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] G080; // 주택계약면적
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(100)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 100)]
//        public char[] G081; // 임대차계약서상주소지
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G082; // 임대차계약기간시작
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(8)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//        public char[] G083; // 임대차계약기간종료
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] G084; // 전세보증금
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] G085; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(186)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 186)]
//        public char[] G086; // 공란
//    }
//    private G_Record G_rec;

//    // / H 기부조정명세 레코드
//    private struct H_record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] H001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] H002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] H003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] H004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] H005; // 사업자번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H006; // 소득자주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] H007; // 내,외국인코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] H008; // 성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] H009; // 유형코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(4)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 4)]
//        public char[] H010; // 기부년도
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H011; // 기부금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H012; // 전년까지공제된금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H013; // 공제대상금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H014; // 해당년도공제금액 필요경비 '0'  2016
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H015; // 해당년도공제금액세액(소득)공제금액  2016
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H016; // 해당년도에공제받지못한금액_소멸금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] H017; // 해당년도에공제받지못한금액_이월금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] H018; // 기부조정명세일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1714)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1714)]
//        public char[] H019; // 공란
//    }
//    private H_record H_rec;

//    // / I 해당년도 기부명세 레코드
//    private struct I_Record
//    {
//        // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] I001; // 레코드구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] I002; // 자료구분
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(3)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 3)]
//        public char[] I003; // 세무서
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(6)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 6)]
//        public char[] I004; // 일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(10)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//        public char[] I005; // 사업자등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I006; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(2)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 2)]
//        public char[] I007; // 유형코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] I008; // 기부내용
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I009; // 기부처-사업자등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(60)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 60)]
//        public char[] I010; // 기부처-법인명(상호)
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] I011; // 관계
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1)]
//        public char[] I012; // 내,외국인코드
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(30)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 30)]
//        public char[] I013; // 성명
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I014; // 주민등록번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] I015; // 건수
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I016; // 기부금합계금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I017; // 공제대상기부금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I018; // 기부장려금신청금액
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(13)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 13)]
//        public char[] I019; // 기타
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(5)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 5)]
//        public char[] I020; // 해당연도기부명세일련번호
//                            // UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//        [VBFixedString(1664)]
//        [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 1664)]
//        public char[] I021; // 공란
//    }
//    private I_Record I_rec;


//    // *******************************************************************
//    // .srf 파일로부터 폼을 로드한다.
//    // *******************************************************************
//    public void LoadForm()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo LoadForm_Error' at character 163028
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


//        oXmlDoc.Load(MDC_Globals.SP_Path + @"\" + SP_Screen + @"\PH_PY980.srf");
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount);

//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        // //여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10);
//        oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10);

//        Sbo_Application.LoadBatchActions((oXmlDoc.xml));

//        oFormUniqueID = "PH_PY980_" + GetTotalFormsCount;

//        // 폼 할당
//        oForm = Sbo_Application.Forms.Item(oFormUniqueID);

//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        // 컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
//        // ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//        AddForms(this, oFormUniqueID, "PH_PY980");
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
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Raise_FormIte...' at character 166038
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
//                        // If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn사원리스트
//                        // oForm.Items("MSTCOD").CLICK ct_Regular
//                        // Sbo_Application.ActivateMenuItem ("7425")
//                        // BubbleEvent = False
//                        // Else

//                        if (pval.ItemUID == "Btn01")
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
//                    }
//                    else
//                    {
//                    }

//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                {
//                    if (pval.BeforeAction == true)
//                    {
//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        if (pval.ItemChanged == true)
//                        {
//                            switch (pval.ItemUID)
//                            {
//                                case "CLTCOD":
//                                    {
//                                        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" + Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        oRecordSet.DoQuery(sQry);
//                                        // UPGRADE_WARNING: oForm.Items(HtaxID).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        oForm.Items.Item("HtaxID").Specific.String = Trim(oRecordSet.Fields.Item("U_HomeTId").VALUE);
//                                        // UPGRADE_WARNING: oForm.Items(TeamName).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        oForm.Items.Item("TeamName").Specific.String = Trim(oRecordSet.Fields.Item("U_ChgDpt").VALUE);
//                                        // UPGRADE_WARNING: oForm.Items(Dname).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        oForm.Items.Item("Dname").Specific.String = Trim(oRecordSet.Fields.Item("U_ChgName").VALUE);
//                                        // UPGRADE_WARNING: oForm.Items(Dtel).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        oForm.Items.Item("Dtel").Specific.String = Trim(oRecordSet.Fields.Item("U_ChgTel").VALUE);
//                                        break;
//                                    }
//                            }
//                        }
//                    }

//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                {
//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_CLICK:
//                {
//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                {
//                    break;
//                }

//            case object _ when SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                {
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
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Raise_FormDat...' at character 171435
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
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 172942
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo Error_Message

// */		SAPbouiCOM.ComboBox oCombo;
//        SAPbobsCOM.Recordset oRecordSet;
//        string sQry;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//        oCombo = oForm.Items.Item("CLTCOD").Specific;
//        oCombo.DataBind.SetBound(true, "", "CLTCOD");
//        oForm.Items.Item("CLTCOD").DisplayDesc = true;
//        // // 접속자에 따른 권한별 사업장 콤보박스세팅
//        CLTCOD_Select(oForm, "CLTCOD");

//        // UPGRADE_WARNING: oForm.Items(YYYY).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("YYYY").Specific.String = System.Convert.ToDouble(VB6.Format(DateTime.Now, "YYYY")) - 1; // 년도 기본년도에서 - 1

//        oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10); // 제출일자
//                                                                                             // UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        Sbo_Application.StatusBar.SetText("CreateItems 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//    }

//    private bool File_Create()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 174918
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

//        // 화면변수를 전역변수로 MOVE
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        CLTCOD = Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        yyyy = Trim(oForm.Items.Item("YYYY").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        HtaxID = Trim(oForm.Items.Item("HtaxID").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        TeamName = Trim(oForm.Items.Item("TeamName").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        Dname = Trim(oForm.Items.Item("Dname").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        Dtel = Trim(oForm.Items.Item("Dtel").Specific.VALUE);
//        // UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        DocDate = Trim(oForm.Items.Item("DocDate").Specific.VALUE);

//        ErrNum = 0;

//        // / Question
//        if (Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }

//        // Sbo_Application.StatusBar.SetText "전산매체수록중..............", bmt_Short, smt_Success

//        // / A RECORD 처리
//        if (File_Create_A_record == false)
//        {
//            ErrNum = 2;
//            goto Error_Message;
//        }

//        // / B RECORD 처리
//        if (File_Create_B_record == false)
//        {
//            ErrNum = 3;
//            goto Error_Message;
//        }

//        // / C RECORD 처리  D.E.F.G 처리
//        if (File_Create_C_record == false)
//        {
//            ErrNum = 4;
//            goto Error_Message;
//        }

//        FileSystem.FileClose(1);

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
//        else
//            Sbo_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        File_Create = false;
//    }
//    private bool File_Create_A_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 179204
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
//        string saup;
//        string CheckA;
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        CheckA = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        // / A_RECORE QUERY
//        sQry = "EXEC PH_PY980_A '" + CLTCOD + "', '" + HtaxID + "', '" + TeamName + "', '" + Dname + "', '" + Dtel + "', '" + DocDate + "'";
//        oRecordSet.DoQuery(sQry);

//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }
//        else
//        {
//            // PATH및 파일이름 만들기
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            saup = oRecordSet.Fields.Item("A009").VALUE; // 사업자번호
//            oFilePath = @"C:\BANK\C" + Strings.Mid(saup, 1, 7) + "." + Strings.Mid(saup, 8, 3);


//            // A RECORD MOVE

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A001 = oRecordSet.Fields.Item("A001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A002 = oRecordSet.Fields.Item("A002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A003 = oRecordSet.Fields.Item("A003").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A004 = oRecordSet.Fields.Item("A004").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A005 = oRecordSet.Fields.Item("A005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A006 = oRecordSet.Fields.Item("A006").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A007 = oRecordSet.Fields.Item("A007").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A008 = oRecordSet.Fields.Item("A008").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A009 = oRecordSet.Fields.Item("A009").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A010 = oRecordSet.Fields.Item("A010").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A011 = oRecordSet.Fields.Item("A011").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A012 = oRecordSet.Fields.Item("A012").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A013 = oRecordSet.Fields.Item("A013").VALUE;

//            A_rec.A014 = VB6.Format(oRecordSet.Fields.Item("A014").VALUE, new string("0", Strings.Len(A_rec.A014)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A015 = oRecordSet.Fields.Item("A015").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            A_rec.A016 = oRecordSet.Fields.Item("A016").VALUE;

//            FileSystem.FileClose(1);
//            FileSystem.FileOpen(1, oFilePath, OpenMode.Output);
//            PrintLine(1, MDC_SetMod.sStr(A_rec.A001) + MDC_SetMod.sStr(A_rec.A002) + MDC_SetMod.sStr(A_rec.A003) + MDC_SetMod.sStr(A_rec.A004) + MDC_SetMod.sStr(A_rec.A005) + MDC_SetMod.sStr(A_rec.A006) + MDC_SetMod.sStr(A_rec.A007) + MDC_SetMod.sStr(A_rec.A008) + MDC_SetMod.sStr(A_rec.A009) + MDC_SetMod.sStr(A_rec.A010) + MDC_SetMod.sStr(A_rec.A011) + MDC_SetMod.sStr(A_rec.A012) + MDC_SetMod.sStr(A_rec.A013) + MDC_SetMod.sStr(A_rec.A014) + MDC_SetMod.sStr(A_rec.A015) + MDC_SetMod.sStr(A_rec.A016));
//        }

//        if (System.Convert.ToBoolean(CheckA) == false)
//            File_Create_A_record = true;
//        else
//            File_Create_A_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("귀속년도의 자사정보(A RECORD)가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("A레코드오류: " + Information.Err.Description, ref false, ref true);

//        File_Create_A_record = false;
//    }

//    private short File_Create_B_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 185636
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

//        // / B_RECORE QUERY
//        sQry = "EXEC PH_PY980_B '" + CLTCOD + "', '" + yyyy + "'";
//        oRecordSet.DoQuery(sQry);

//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }
//        else
//        {
//            // B RECORD MOVE

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B001 = oRecordSet.Fields.Item("B001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B002 = oRecordSet.Fields.Item("B002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B003 = oRecordSet.Fields.Item("B003").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B004 = oRecordSet.Fields.Item("B004").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B005 = oRecordSet.Fields.Item("B005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B006 = oRecordSet.Fields.Item("B006").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B007 = oRecordSet.Fields.Item("B007").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B008 = oRecordSet.Fields.Item("B008").VALUE;
//            B_rec.B009 = VB6.Format(oRecordSet.Fields.Item("B009").VALUE, new string("0", Strings.Len(B_rec.B009)));
//            B_rec.B010 = VB6.Format(oRecordSet.Fields.Item("B010").VALUE, new string("0", Strings.Len(B_rec.B010)));
//            B_rec.B011 = VB6.Format(oRecordSet.Fields.Item("B011").VALUE, new string("0", Strings.Len(B_rec.B011)));
//            B_rec.B012 = VB6.Format(oRecordSet.Fields.Item("B012").VALUE, new string("0", Strings.Len(B_rec.B012)));
//            B_rec.B013 = VB6.Format(oRecordSet.Fields.Item("B013").VALUE, new string("0", Strings.Len(B_rec.B013)));
//            B_rec.B014 = VB6.Format(oRecordSet.Fields.Item("B014").VALUE, new string("0", Strings.Len(B_rec.B014)));
//            B_rec.B015 = VB6.Format(oRecordSet.Fields.Item("B015").VALUE, new string("0", Strings.Len(B_rec.B015)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B016 = oRecordSet.Fields.Item("B016").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            B_rec.B017 = oRecordSet.Fields.Item("B017").VALUE;

//            // Close #1
//            // Open oFilePath For Output As #1
//            PrintLine(1, MDC_SetMod.sStr(B_rec.B001) + MDC_SetMod.sStr(B_rec.B002) + MDC_SetMod.sStr(B_rec.B003) + MDC_SetMod.sStr(B_rec.B004) + MDC_SetMod.sStr(B_rec.B005) + MDC_SetMod.sStr(B_rec.B006) + MDC_SetMod.sStr(B_rec.B007) + MDC_SetMod.sStr(B_rec.B008) + MDC_SetMod.sStr(B_rec.B009) + MDC_SetMod.sStr(B_rec.B010) + MDC_SetMod.sStr(B_rec.B011) + MDC_SetMod.sStr(B_rec.B012) + MDC_SetMod.sStr(B_rec.B013) + MDC_SetMod.sStr(B_rec.B014) + MDC_SetMod.sStr(B_rec.B015) + MDC_SetMod.sStr(B_rec.B016) + MDC_SetMod.sStr(B_rec.B017));
//        }

//        if (System.Convert.ToBoolean(CheckB) == false)
//            File_Create_B_record = true;
//        else
//            File_Create_B_record = false;

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
//            Sbo_Application.StatusBar.SetText("B레코드가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            File_Create_B_record = 1;
//        }
//        else
//        {
//            Matrix_AddRow("B레코드오류: " + Information.Err.Description, ref false);
//            File_Create_B_record = 1;
//        }
//    }

//    private bool File_Create_C_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 190959
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

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        CheckC = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        // / C_RECORE QUERY
//        sQry = "EXEC PH_PY980_C '" + CLTCOD + "', '" + yyyy + "'";

//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }

//        SAPbouiCOM.ProgressBar ProgressBar01;
//        ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, false);

//        NEWCNT = 0;
//        while (!oRecordSet.EOF)
//        {
//            NEWCNT = NEWCNT + 1; // / 일련번호

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_SAUP = oRecordSet.Fields.Item("saup").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_YYYY = oRecordSet.Fields.Item("yyyy").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_SABUN = oRecordSet.Fields.Item("sabun").VALUE;

//            // C RECORD MOVE

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C001 = oRecordSet.Fields.Item("C001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C002 = oRecordSet.Fields.Item("C002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C003 = oRecordSet.Fields.Item("C003").VALUE;
//            C_rec.C004 = VB6.Format(NEWCNT, new string("0", Strings.Len(C_rec.C004))); // / 일련번호
//                                                                                       // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C005 = oRecordSet.Fields.Item("C005").VALUE;
//            C_rec.C006 = VB6.Format(oRecordSet.Fields.Item("C006").VALUE, new string("0", Strings.Len(C_rec.C006)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C007 = oRecordSet.Fields.Item("C007").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C008 = oRecordSet.Fields.Item("C008").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C009 = oRecordSet.Fields.Item("C009").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C010 = oRecordSet.Fields.Item("C010").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C011 = oRecordSet.Fields.Item("C011").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C012 = oRecordSet.Fields.Item("C012").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C013 = oRecordSet.Fields.Item("C013").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C014 = oRecordSet.Fields.Item("C014").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C015 = oRecordSet.Fields.Item("C015").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C016 = oRecordSet.Fields.Item("C016").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C017 = oRecordSet.Fields.Item("C017").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C018 = oRecordSet.Fields.Item("C018").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C019 = oRecordSet.Fields.Item("C019").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C020 = oRecordSet.Fields.Item("C020").VALUE;
//            C_rec.C021 = VB6.Format(oRecordSet.Fields.Item("C021").VALUE, new string("0", Strings.Len(C_rec.C021)));
//            C_rec.C022 = VB6.Format(oRecordSet.Fields.Item("C022").VALUE, new string("0", Strings.Len(C_rec.C022)));
//            C_rec.C023 = VB6.Format(oRecordSet.Fields.Item("C023").VALUE, new string("0", Strings.Len(C_rec.C023)));
//            C_rec.C024 = VB6.Format(oRecordSet.Fields.Item("C024").VALUE, new string("0", Strings.Len(C_rec.C024)));
//            C_rec.C025 = VB6.Format(oRecordSet.Fields.Item("C025").VALUE, new string("0", Strings.Len(C_rec.C025)));
//            C_rec.C026 = VB6.Format(oRecordSet.Fields.Item("C026").VALUE, new string("0", Strings.Len(C_rec.C026)));
//            C_rec.C027 = VB6.Format(oRecordSet.Fields.Item("C027").VALUE, new string("0", Strings.Len(C_rec.C027)));
//            C_rec.C028 = VB6.Format(oRecordSet.Fields.Item("C028").VALUE, new string("0", Strings.Len(C_rec.C028)));
//            C_rec.C029 = VB6.Format(oRecordSet.Fields.Item("C029").VALUE, new string("0", Strings.Len(C_rec.C029)));
//            C_rec.C030 = VB6.Format(oRecordSet.Fields.Item("C030").VALUE, new string("0", Strings.Len(C_rec.C030)));
//            C_rec.C031 = VB6.Format(oRecordSet.Fields.Item("C031").VALUE, new string("0", Strings.Len(C_rec.C031)));
//            C_rec.C032 = VB6.Format(oRecordSet.Fields.Item("C032").VALUE, new string("0", Strings.Len(C_rec.C032)));

//            C_rec.C034 = VB6.Format(oRecordSet.Fields.Item("C034").VALUE, new string("0", Strings.Len(C_rec.C034)));
//            C_rec.C035 = VB6.Format(oRecordSet.Fields.Item("C035").VALUE, new string("0", Strings.Len(C_rec.C035)));
//            C_rec.C036 = VB6.Format(oRecordSet.Fields.Item("C036").VALUE, new string("0", Strings.Len(C_rec.C036)));
//            C_rec.C037 = VB6.Format(oRecordSet.Fields.Item("C037").VALUE, new string("0", Strings.Len(C_rec.C037)));
//            C_rec.C038 = VB6.Format(oRecordSet.Fields.Item("C038").VALUE, new string("0", Strings.Len(C_rec.C038)));
//            C_rec.C039 = VB6.Format(oRecordSet.Fields.Item("C039").VALUE, new string("0", Strings.Len(C_rec.C039)));
//            C_rec.C040 = VB6.Format(oRecordSet.Fields.Item("C040").VALUE, new string("0", Strings.Len(C_rec.C040)));
//            C_rec.C041 = VB6.Format(oRecordSet.Fields.Item("C041").VALUE, new string("0", Strings.Len(C_rec.C041)));
//            C_rec.C042 = VB6.Format(oRecordSet.Fields.Item("C042").VALUE, new string("0", Strings.Len(C_rec.C042)));
//            C_rec.C043 = VB6.Format(oRecordSet.Fields.Item("C043").VALUE, new string("0", Strings.Len(C_rec.C043)));
//            C_rec.C044 = VB6.Format(oRecordSet.Fields.Item("C044").VALUE, new string("0", Strings.Len(C_rec.C044)));
//            C_rec.C045 = VB6.Format(oRecordSet.Fields.Item("C045").VALUE, new string("0", Strings.Len(C_rec.C045)));
//            C_rec.C046 = VB6.Format(oRecordSet.Fields.Item("C046").VALUE, new string("0", Strings.Len(C_rec.C046)));
//            C_rec.C047 = VB6.Format(oRecordSet.Fields.Item("C047").VALUE, new string("0", Strings.Len(C_rec.C047)));
//            C_rec.C048 = VB6.Format(oRecordSet.Fields.Item("C048").VALUE, new string("0", Strings.Len(C_rec.C048)));
//            C_rec.C049 = VB6.Format(oRecordSet.Fields.Item("C049").VALUE, new string("0", Strings.Len(C_rec.C049)));
//            C_rec.C050 = VB6.Format(oRecordSet.Fields.Item("C050").VALUE, new string("0", Strings.Len(C_rec.C050)));
//            C_rec.C051 = VB6.Format(oRecordSet.Fields.Item("C051").VALUE, new string("0", Strings.Len(C_rec.C051)));
//            C_rec.C052 = VB6.Format(oRecordSet.Fields.Item("C052").VALUE, new string("0", Strings.Len(C_rec.C052)));
//            C_rec.C053 = VB6.Format(oRecordSet.Fields.Item("C053").VALUE, new string("0", Strings.Len(C_rec.C053)));
//            C_rec.C054 = VB6.Format(oRecordSet.Fields.Item("C054").VALUE, new string("0", Strings.Len(C_rec.C054)));
//            C_rec.C055 = VB6.Format(oRecordSet.Fields.Item("C055").VALUE, new string("0", Strings.Len(C_rec.C055)));
//            C_rec.C056 = VB6.Format(oRecordSet.Fields.Item("C056").VALUE, new string("0", Strings.Len(C_rec.C056)));
//            C_rec.C057 = VB6.Format(oRecordSet.Fields.Item("C057").VALUE, new string("0", Strings.Len(C_rec.C057)));
//            C_rec.C058 = VB6.Format(oRecordSet.Fields.Item("C058").VALUE, new string("0", Strings.Len(C_rec.C058)));
//            C_rec.C059 = VB6.Format(oRecordSet.Fields.Item("C059").VALUE, new string("0", Strings.Len(C_rec.C059)));
//            C_rec.C060 = VB6.Format(oRecordSet.Fields.Item("C060").VALUE, new string("0", Strings.Len(C_rec.C060)));
//            C_rec.C061A = VB6.Format(oRecordSet.Fields.Item("C061A").VALUE, new string("0", Strings.Len(C_rec.C061A)));
//            C_rec.C061B = VB6.Format(oRecordSet.Fields.Item("C061B").VALUE, new string("0", Strings.Len(C_rec.C061B)));
//            C_rec.C061C = VB6.Format(oRecordSet.Fields.Item("C061C").VALUE, new string("0", Strings.Len(C_rec.C061C)));
//            C_rec.C062 = VB6.Format(oRecordSet.Fields.Item("C062").VALUE, new string("0", Strings.Len(C_rec.C062)));
//            C_rec.C063 = VB6.Format(oRecordSet.Fields.Item("C063").VALUE, new string("0", Strings.Len(C_rec.C063)));
//            C_rec.C064A = VB6.Format(oRecordSet.Fields.Item("C064A").VALUE, new string("0", Strings.Len(C_rec.C064A)));
//            C_rec.C064B = VB6.Format(oRecordSet.Fields.Item("C064B").VALUE, new string("0", Strings.Len(C_rec.C064B)));
//            C_rec.C064C = VB6.Format(oRecordSet.Fields.Item("C064C").VALUE, new string("0", Strings.Len(C_rec.C064C)));
//            C_rec.C064D = VB6.Format(oRecordSet.Fields.Item("C064D").VALUE, new string("0", Strings.Len(C_rec.C064D)));
//            C_rec.C065 = VB6.Format(oRecordSet.Fields.Item("C065").VALUE, new string("0", Strings.Len(C_rec.C065)));
//            C_rec.C066 = VB6.Format(oRecordSet.Fields.Item("C066").VALUE, new string("0", Strings.Len(C_rec.C066)));

//            C_rec.C068 = VB6.Format(oRecordSet.Fields.Item("C068").VALUE, new string("0", Strings.Len(C_rec.C068)));
//            C_rec.C069 = VB6.Format(oRecordSet.Fields.Item("C069").VALUE, new string("0", Strings.Len(C_rec.C069)));
//            C_rec.C070 = VB6.Format(oRecordSet.Fields.Item("C070").VALUE, new string("0", Strings.Len(C_rec.C070)));
//            C_rec.C071 = VB6.Format(oRecordSet.Fields.Item("C071").VALUE, new string("0", Strings.Len(C_rec.C071)));
//            C_rec.C072 = VB6.Format(oRecordSet.Fields.Item("C072").VALUE, new string("0", Strings.Len(C_rec.C072)));
//            C_rec.C073 = VB6.Format(oRecordSet.Fields.Item("C073").VALUE, new string("0", Strings.Len(C_rec.C073)));
//            C_rec.C074 = VB6.Format(oRecordSet.Fields.Item("C074").VALUE, new string("0", Strings.Len(C_rec.C074)));
//            C_rec.C075A = VB6.Format(oRecordSet.Fields.Item("C075A").VALUE, new string("0", Strings.Len(C_rec.C075A)));
//            C_rec.C075B = VB6.Format(oRecordSet.Fields.Item("C075B").VALUE, new string("0", Strings.Len(C_rec.C075B)));
//            C_rec.C076A = VB6.Format(oRecordSet.Fields.Item("C076A").VALUE, new string("0", Strings.Len(C_rec.C076A)));
//            C_rec.C076B = VB6.Format(oRecordSet.Fields.Item("C076B").VALUE, new string("0", Strings.Len(C_rec.C076B)));
//            C_rec.C077A = VB6.Format(oRecordSet.Fields.Item("C077A").VALUE, new string("0", Strings.Len(C_rec.C077A)));
//            C_rec.C077B = VB6.Format(oRecordSet.Fields.Item("C077B").VALUE, new string("0", Strings.Len(C_rec.C077B)));
//            C_rec.C078 = VB6.Format(oRecordSet.Fields.Item("C078").VALUE, new string("0", Strings.Len(C_rec.C078)));
//            C_rec.C079 = VB6.Format(oRecordSet.Fields.Item("C079").VALUE, new string("0", Strings.Len(C_rec.C079)));
//            C_rec.C080A = VB6.Format(oRecordSet.Fields.Item("C080A").VALUE, new string("0", Strings.Len(C_rec.C080A)));
//            C_rec.C080B = VB6.Format(oRecordSet.Fields.Item("C080B").VALUE, new string("0", Strings.Len(C_rec.C080B)));
//            C_rec.C081A = VB6.Format(oRecordSet.Fields.Item("C081A").VALUE, new string("0", Strings.Len(C_rec.C081A)));
//            C_rec.C081B = VB6.Format(oRecordSet.Fields.Item("C081B").VALUE, new string("0", Strings.Len(C_rec.C081B)));
//            C_rec.C082A = VB6.Format(oRecordSet.Fields.Item("C082A").VALUE, new string("0", Strings.Len(C_rec.C082A)));
//            C_rec.C082B = VB6.Format(oRecordSet.Fields.Item("C082B").VALUE, new string("0", Strings.Len(C_rec.C082B)));
//            C_rec.C083A = VB6.Format(oRecordSet.Fields.Item("C083A").VALUE, new string("0", Strings.Len(C_rec.C083A)));
//            C_rec.C083B = VB6.Format(oRecordSet.Fields.Item("C083B").VALUE, new string("0", Strings.Len(C_rec.C083B)));
//            C_rec.C084A = VB6.Format(oRecordSet.Fields.Item("C084A").VALUE, new string("0", Strings.Len(C_rec.C084A)));
//            C_rec.C084B = VB6.Format(oRecordSet.Fields.Item("C084B").VALUE, new string("0", Strings.Len(C_rec.C084B)));
//            C_rec.C085A = VB6.Format(oRecordSet.Fields.Item("C085A").VALUE, new string("0", Strings.Len(C_rec.C085A)));
//            C_rec.C085B = VB6.Format(oRecordSet.Fields.Item("C085B").VALUE, new string("0", Strings.Len(C_rec.C085B)));
//            C_rec.C086A = VB6.Format(oRecordSet.Fields.Item("C086A").VALUE, new string("0", Strings.Len(C_rec.C086A)));
//            C_rec.C086B = VB6.Format(oRecordSet.Fields.Item("C086B").VALUE, new string("0", Strings.Len(C_rec.C086B)));
//            C_rec.C087A = VB6.Format(oRecordSet.Fields.Item("C087A").VALUE, new string("0", Strings.Len(C_rec.C087A)));
//            C_rec.C087B = VB6.Format(oRecordSet.Fields.Item("C087B").VALUE, new string("0", Strings.Len(C_rec.C087B)));
//            C_rec.C088A = VB6.Format(oRecordSet.Fields.Item("C088A").VALUE, new string("0", Strings.Len(C_rec.C088A)));
//            C_rec.C088B = VB6.Format(oRecordSet.Fields.Item("C088B").VALUE, new string("0", Strings.Len(C_rec.C088B)));
//            C_rec.C088C = VB6.Format(oRecordSet.Fields.Item("C088C").VALUE, new string("0", Strings.Len(C_rec.C088C)));
//            C_rec.C089A = VB6.Format(oRecordSet.Fields.Item("C089A").VALUE, new string("0", Strings.Len(C_rec.C089A)));
//            C_rec.C089B = VB6.Format(oRecordSet.Fields.Item("C089B").VALUE, new string("0", Strings.Len(C_rec.C089B)));
//            C_rec.C090A = VB6.Format(oRecordSet.Fields.Item("C090A").VALUE, new string("0", Strings.Len(C_rec.C090A)));
//            C_rec.C090B = VB6.Format(oRecordSet.Fields.Item("C090B").VALUE, new string("0", Strings.Len(C_rec.C090B)));
//            C_rec.C090C = VB6.Format(oRecordSet.Fields.Item("C090C").VALUE, new string("0", Strings.Len(C_rec.C090C)));
//            C_rec.C090D = VB6.Format(oRecordSet.Fields.Item("C090D").VALUE, new string("0", Strings.Len(C_rec.C090D)));
//            C_rec.C091 = VB6.Format(oRecordSet.Fields.Item("C091").VALUE, new string("0", Strings.Len(C_rec.C091)));
//            C_rec.C092 = VB6.Format(oRecordSet.Fields.Item("C092").VALUE, new string("0", Strings.Len(C_rec.C092)));

//            C_rec.C094 = VB6.Format(oRecordSet.Fields.Item("C094").VALUE, new string("0", Strings.Len(C_rec.C094)));
//            C_rec.C095 = VB6.Format(oRecordSet.Fields.Item("C095").VALUE, new string("0", Strings.Len(C_rec.C095)));
//            C_rec.C096 = VB6.Format(oRecordSet.Fields.Item("C096").VALUE, new string("0", Strings.Len(C_rec.C096)));
//            C_rec.C097 = VB6.Format(oRecordSet.Fields.Item("C097").VALUE, new string("0", Strings.Len(C_rec.C097)));
//            C_rec.C098 = VB6.Format(oRecordSet.Fields.Item("C098").VALUE, new string("0", Strings.Len(C_rec.C098)));
//            C_rec.C099 = VB6.Format(oRecordSet.Fields.Item("C099").VALUE, new string("0", Strings.Len(C_rec.C099)));
//            C_rec.C100 = VB6.Format(oRecordSet.Fields.Item("C100").VALUE, new string("0", Strings.Len(C_rec.C100)));
//            C_rec.C101 = VB6.Format(oRecordSet.Fields.Item("C101").VALUE, new string("0", Strings.Len(C_rec.C101)));
//            C_rec.C102 = VB6.Format(oRecordSet.Fields.Item("C102").VALUE, new string("0", Strings.Len(C_rec.C102)));
//            C_rec.C103 = VB6.Format(oRecordSet.Fields.Item("C103").VALUE, new string("0", Strings.Len(C_rec.C103)));
//            C_rec.C104 = VB6.Format(oRecordSet.Fields.Item("C104").VALUE, new string("0", Strings.Len(C_rec.C104)));
//            C_rec.C105 = VB6.Format(oRecordSet.Fields.Item("C105").VALUE, new string("0", Strings.Len(C_rec.C105)));
//            C_rec.C106 = VB6.Format(oRecordSet.Fields.Item("C106").VALUE, new string("0", Strings.Len(C_rec.C106)));

//            C_rec.C108 = VB6.Format(oRecordSet.Fields.Item("C108").VALUE, new string("0", Strings.Len(C_rec.C108)));
//            C_rec.C109 = VB6.Format(oRecordSet.Fields.Item("C109").VALUE, new string("0", Strings.Len(C_rec.C109)));
//            C_rec.C110 = VB6.Format(oRecordSet.Fields.Item("C110").VALUE, new string("0", Strings.Len(C_rec.C110)));
//            C_rec.C111 = VB6.Format(oRecordSet.Fields.Item("C111").VALUE, new string("0", Strings.Len(C_rec.C111)));
//            C_rec.C112 = VB6.Format(oRecordSet.Fields.Item("C112").VALUE, new string("0", Strings.Len(C_rec.C112)));
//            C_rec.C113 = VB6.Format(oRecordSet.Fields.Item("C113").VALUE, new string("0", Strings.Len(C_rec.C113)));
//            C_rec.C114 = VB6.Format(oRecordSet.Fields.Item("C114").VALUE, new string("0", Strings.Len(C_rec.C114)));
//            C_rec.C115 = VB6.Format(oRecordSet.Fields.Item("C115").VALUE, new string("0", Strings.Len(C_rec.C115)));
//            C_rec.C116 = VB6.Format(oRecordSet.Fields.Item("C116").VALUE, new string("0", Strings.Len(C_rec.C116)));

//            C_rec.C118 = VB6.Format(oRecordSet.Fields.Item("C118").VALUE, new string("0", Strings.Len(C_rec.C118)));
//            C_rec.C119 = VB6.Format(oRecordSet.Fields.Item("C119").VALUE, new string("0", Strings.Len(C_rec.C119)));
//            C_rec.C120A = VB6.Format(oRecordSet.Fields.Item("C120A").VALUE, new string("0", Strings.Len(C_rec.C120A)));
//            C_rec.C120B = VB6.Format(oRecordSet.Fields.Item("C120B").VALUE, new string("0", Strings.Len(C_rec.C120B)));
//            C_rec.C121A = VB6.Format(oRecordSet.Fields.Item("C121A").VALUE, new string("0", Strings.Len(C_rec.C121A)));
//            C_rec.C121B = VB6.Format(oRecordSet.Fields.Item("C121B").VALUE, new string("0", Strings.Len(C_rec.C121B)));
//            C_rec.C122A = VB6.Format(oRecordSet.Fields.Item("C122A").VALUE, new string("0", Strings.Len(C_rec.C122A)));
//            C_rec.C122B = VB6.Format(oRecordSet.Fields.Item("C122B").VALUE, new string("0", Strings.Len(C_rec.C122B)));
//            C_rec.C123A = VB6.Format(oRecordSet.Fields.Item("C123A").VALUE, new string("0", Strings.Len(C_rec.C123A)));
//            C_rec.C123B = VB6.Format(oRecordSet.Fields.Item("C123B").VALUE, new string("0", Strings.Len(C_rec.C123B)));
//            C_rec.C124A = VB6.Format(oRecordSet.Fields.Item("C124A").VALUE, new string("0", Strings.Len(C_rec.C124A)));
//            C_rec.C124B = VB6.Format(oRecordSet.Fields.Item("C124B").VALUE, new string("0", Strings.Len(C_rec.C124B)));
//            C_rec.C125A = VB6.Format(oRecordSet.Fields.Item("C125A").VALUE, new string("0", Strings.Len(C_rec.C125A)));
//            C_rec.C125B = VB6.Format(oRecordSet.Fields.Item("C125B").VALUE, new string("0", Strings.Len(C_rec.C125B)));
//            C_rec.C126A = VB6.Format(oRecordSet.Fields.Item("C126A").VALUE, new string("0", Strings.Len(C_rec.C126A)));
//            C_rec.C126B = VB6.Format(oRecordSet.Fields.Item("C126B").VALUE, new string("0", Strings.Len(C_rec.C126B)));
//            C_rec.C127A = VB6.Format(oRecordSet.Fields.Item("C127A").VALUE, new string("0", Strings.Len(C_rec.C127A)));
//            C_rec.C127B = VB6.Format(oRecordSet.Fields.Item("C127B").VALUE, new string("0", Strings.Len(C_rec.C127B)));
//            C_rec.C128A = VB6.Format(oRecordSet.Fields.Item("C128A").VALUE, new string("0", Strings.Len(C_rec.C128A)));
//            C_rec.C128B = VB6.Format(oRecordSet.Fields.Item("C128B").VALUE, new string("0", Strings.Len(C_rec.C128B)));
//            C_rec.C129A = VB6.Format(oRecordSet.Fields.Item("C129A").VALUE, new string("0", Strings.Len(C_rec.C129A)));
//            C_rec.C129B = VB6.Format(oRecordSet.Fields.Item("C129B").VALUE, new string("0", Strings.Len(C_rec.C129B)));
//            C_rec.C130A = VB6.Format(oRecordSet.Fields.Item("C130A").VALUE, new string("0", Strings.Len(C_rec.C130A)));
//            C_rec.C130B = VB6.Format(oRecordSet.Fields.Item("C130B").VALUE, new string("0", Strings.Len(C_rec.C130B)));
//            C_rec.C131A = VB6.Format(oRecordSet.Fields.Item("C131A").VALUE, new string("0", Strings.Len(C_rec.C131A)));
//            C_rec.C131B = VB6.Format(oRecordSet.Fields.Item("C131B").VALUE, new string("0", Strings.Len(C_rec.C131B)));
//            C_rec.C132A = VB6.Format(oRecordSet.Fields.Item("C132A").VALUE, new string("0", Strings.Len(C_rec.C132A)));
//            C_rec.C132B = VB6.Format(oRecordSet.Fields.Item("C132B").VALUE, new string("0", Strings.Len(C_rec.C132B)));
//            C_rec.C133A = VB6.Format(oRecordSet.Fields.Item("C133A").VALUE, new string("0", Strings.Len(C_rec.C133A)));
//            C_rec.C133B = VB6.Format(oRecordSet.Fields.Item("C133B").VALUE, new string("0", Strings.Len(C_rec.C133B)));
//            C_rec.C134A = VB6.Format(oRecordSet.Fields.Item("C134A").VALUE, new string("0", Strings.Len(C_rec.C134A)));
//            C_rec.C134B = VB6.Format(oRecordSet.Fields.Item("C134B").VALUE, new string("0", Strings.Len(C_rec.C134B)));
//            C_rec.C135 = VB6.Format(oRecordSet.Fields.Item("C135").VALUE, new string("0", Strings.Len(C_rec.C135)));

//            C_rec.C137 = VB6.Format(oRecordSet.Fields.Item("C137").VALUE, new string("0", Strings.Len(C_rec.C137)));
//            C_rec.C138 = VB6.Format(oRecordSet.Fields.Item("C138").VALUE, new string("0", Strings.Len(C_rec.C138)));
//            C_rec.C139 = VB6.Format(oRecordSet.Fields.Item("C139").VALUE, new string("0", Strings.Len(C_rec.C139)));
//            C_rec.C140 = VB6.Format(oRecordSet.Fields.Item("C140").VALUE, new string("0", Strings.Len(C_rec.C140)));
//            C_rec.C141 = VB6.Format(oRecordSet.Fields.Item("C141").VALUE, new string("0", Strings.Len(C_rec.C141)));
//            C_rec.C142A = VB6.Format(oRecordSet.Fields.Item("C142A").VALUE, new string("0", Strings.Len(C_rec.C142A)));
//            C_rec.C142B = VB6.Format(oRecordSet.Fields.Item("C142B").VALUE, new string("0", Strings.Len(C_rec.C142B)));
//            C_rec.C143 = VB6.Format(oRecordSet.Fields.Item("C143").VALUE, new string("0", Strings.Len(C_rec.C143)));

//            C_rec.C145 = VB6.Format(oRecordSet.Fields.Item("C145").VALUE, new string("0", Strings.Len(C_rec.C145)));
//            C_rec.C146A = VB6.Format(oRecordSet.Fields.Item("C146A").VALUE, new string("0", Strings.Len(C_rec.C146A)));
//            C_rec.C146B = VB6.Format(oRecordSet.Fields.Item("C146B").VALUE, new string("0", Strings.Len(C_rec.C146B)));
//            C_rec.C146C = VB6.Format(oRecordSet.Fields.Item("C146C").VALUE, new string("0", Strings.Len(C_rec.C146C)));
//            C_rec.C147A = VB6.Format(oRecordSet.Fields.Item("C147A").VALUE, new string("0", Strings.Len(C_rec.C147A)));
//            C_rec.C147B = VB6.Format(oRecordSet.Fields.Item("C147B").VALUE, new string("0", Strings.Len(C_rec.C147B)));
//            C_rec.C147C = VB6.Format(oRecordSet.Fields.Item("C147C").VALUE, new string("0", Strings.Len(C_rec.C147C)));
//            C_rec.C148A = VB6.Format(oRecordSet.Fields.Item("C148A").VALUE, new string("0", Strings.Len(C_rec.C148A)));
//            C_rec.C148B = VB6.Format(oRecordSet.Fields.Item("C148B").VALUE, new string("0", Strings.Len(C_rec.C148B)));
//            C_rec.C148C = VB6.Format(oRecordSet.Fields.Item("C148C").VALUE, new string("0", Strings.Len(C_rec.C148C)));
//            C_rec.C149A_1 = VB6.Format(oRecordSet.Fields.Item("C149A_1").VALUE, new string("0", Strings.Len(C_rec.C149A_1)));
//            C_rec.C149A_2 = VB6.Format(oRecordSet.Fields.Item("C149A_2").VALUE, new string("0", Strings.Len(C_rec.C149A_2)));
//            C_rec.C149B_1 = VB6.Format(oRecordSet.Fields.Item("C149B_1").VALUE, new string("0", Strings.Len(C_rec.C149B_1)));
//            C_rec.C149B_2 = VB6.Format(oRecordSet.Fields.Item("C149B_2").VALUE, new string("0", Strings.Len(C_rec.C149B_2)));
//            C_rec.C149C_1 = VB6.Format(oRecordSet.Fields.Item("C149C_1").VALUE, new string("0", Strings.Len(C_rec.C149C_1)));
//            C_rec.C149C_2 = VB6.Format(oRecordSet.Fields.Item("C149C_2").VALUE, new string("0", Strings.Len(C_rec.C149C_2)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            C_rec.C150 = oRecordSet.Fields.Item("C150").VALUE;

//            // 예제
//            // C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").VALUE, "-", "")

//            // OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").VALUE) + Val(oRecordSet.Fields("U_BIGWA3").VALUE) + Val(oRecordSet.Fields("U_BIGWA5").VALUE) _
//            // '        + Val(oRecordSet.Fields("U_BIGWA6").VALUE) + Val(oRecordSet.Fields("U_BIGWU3").VALUE)

//            // C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
//            // C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
//            // C_rec.FILLER = Space$(Len(C_rec.FILLER))
//            // C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, , String$(Len(C_rec.C022), "0"))


//            PrintLine(1, MDC_SetMod.sStr(C_rec.C001) + MDC_SetMod.sStr(C_rec.C002) + MDC_SetMod.sStr(C_rec.C003) + MDC_SetMod.sStr(C_rec.C004) + MDC_SetMod.sStr(C_rec.C005) + MDC_SetMod.sStr(C_rec.C006) + MDC_SetMod.sStr(C_rec.C007) + MDC_SetMod.sStr(C_rec.C008) + MDC_SetMod.sStr(C_rec.C009) + MDC_SetMod.sStr(C_rec.C010) + MDC_SetMod.sStr(C_rec.C011) + MDC_SetMod.sStr(C_rec.C012) + MDC_SetMod.sStr(C_rec.C013) + MDC_SetMod.sStr(C_rec.C014) + MDC_SetMod.sStr(C_rec.C015) + MDC_SetMod.sStr(C_rec.C016) + MDC_SetMod.sStr(C_rec.C017) + MDC_SetMod.sStr(C_rec.C018) + MDC_SetMod.sStr(C_rec.C019) + MDC_SetMod.sStr(C_rec.C020) + MDC_SetMod.sStr(C_rec.C021) + MDC_SetMod.sStr(C_rec.C022) + MDC_SetMod.sStr(C_rec.C023) + MDC_SetMod.sStr(C_rec.C024) + MDC_SetMod.sStr(C_rec.C025) + MDC_SetMod.sStr(C_rec.C026) + MDC_SetMod.sStr(C_rec.C027) + MDC_SetMod.sStr(C_rec.C028) + MDC_SetMod.sStr(C_rec.C029) + MDC_SetMod.sStr(C_rec.C030) + MDC_SetMod.sStr(C_rec.C031) + MDC_SetMod.sStr(C_rec.C032) + MDC_SetMod.sStr(C_rec.C034) + MDC_SetMod.sStr(C_rec.C035) + MDC_SetMod.sStr(C_rec.C036) + MDC_SetMod.sStr(C_rec.C037) + MDC_SetMod.sStr(C_rec.C038) + MDC_SetMod.sStr(C_rec.C039) + MDC_SetMod.sStr(C_rec.C040) + MDC_SetMod.sStr(C_rec.C041) + MDC_SetMod.sStr(C_rec.C042) + MDC_SetMod.sStr(C_rec.C043) + MDC_SetMod.sStr(C_rec.C044) + MDC_SetMod.sStr(C_rec.C045) + MDC_SetMod.sStr(C_rec.C046) + MDC_SetMod.sStr(C_rec.C047) + MDC_SetMod.sStr(C_rec.C048) + MDC_SetMod.sStr(C_rec.C049) + MDC_SetMod.sStr(C_rec.C050) + MDC_SetMod.sStr(C_rec.C051) + MDC_SetMod.sStr(C_rec.C052) + MDC_SetMod.sStr(C_rec.C053) + MDC_SetMod.sStr(C_rec.C054) + MDC_SetMod.sStr(C_rec.C055) + MDC_SetMod.sStr(C_rec.C056) + MDC_SetMod.sStr(C_rec.C057) + MDC_SetMod.sStr(C_rec.C058) + MDC_SetMod.sStr(C_rec.C059) + MDC_SetMod.sStr(C_rec.C060) + MDC_SetMod.sStr(C_rec.C061A) + MDC_SetMod.sStr(C_rec.C061B) + MDC_SetMod.sStr(C_rec.C061C) + MDC_SetMod.sStr(C_rec.C062) + MDC_SetMod.sStr(C_rec.C063) + MDC_SetMod.sStr(C_rec.C064A) + MDC_SetMod.sStr(C_rec.C064B) + MDC_SetMod.sStr(C_rec.C064C) + MDC_SetMod.sStr(C_rec.C064D) + MDC_SetMod.sStr(C_rec.C065) + MDC_SetMod.sStr(C_rec.C066) + MDC_SetMod.sStr(C_rec.C068) + MDC_SetMod.sStr(C_rec.C069) + MDC_SetMod.sStr(C_rec.C070) + MDC_SetMod.sStr(C_rec.C071) + MDC_SetMod.sStr(C_rec.C072) + MDC_SetMod.sStr(C_rec.C073) + MDC_SetMod.sStr(C_rec.C074) + MDC_SetMod.sStr(C_rec.C075A) + MDC_SetMod.sStr(C_rec.C075B) + MDC_SetMod.sStr(C_rec.C076A) + MDC_SetMod.sStr(C_rec.C076B) + MDC_SetMod.sStr(C_rec.C077A) + MDC_SetMod.sStr(C_rec.C077B) + MDC_SetMod.sStr(C_rec.C078) + MDC_SetMod.sStr(C_rec.C079) + MDC_SetMod.sStr(C_rec.C080A) + MDC_SetMod.sStr(C_rec.C080B) + MDC_SetMod.sStr(C_rec.C081A) + MDC_SetMod.sStr(C_rec.C081B) + MDC_SetMod.sStr(C_rec.C082A) + MDC_SetMod.sStr(C_rec.C082B) + MDC_SetMod.sStr(C_rec.C083A) + MDC_SetMod.sStr(C_rec.C083B) + MDC_SetMod.sStr(C_rec.C084A) + MDC_SetMod.sStr(C_rec.C084B) + MDC_SetMod.sStr(C_rec.C085A) + MDC_SetMod.sStr(C_rec.C085B) + MDC_SetMod.sStr(C_rec.C086A) + MDC_SetMod.sStr(C_rec.C086B) + MDC_SetMod.sStr(C_rec.C087A) + MDC_SetMod.sStr(C_rec.C087B) + MDC_SetMod.sStr(C_rec.C088A) + MDC_SetMod.sStr(C_rec.C088B) + MDC_SetMod.sStr(C_rec.C088C) + MDC_SetMod.sStr(C_rec.C089A) + MDC_SetMod.sStr(C_rec.C089B) + MDC_SetMod.sStr(C_rec.C090A) + MDC_SetMod.sStr(C_rec.C090B) + MDC_SetMod.sStr(C_rec.C090C) + MDC_SetMod.sStr(C_rec.C090D) + MDC_SetMod.sStr(C_rec.C091) + MDC_SetMod.sStr(C_rec.C092) + MDC_SetMod.sStr(C_rec.C094) + MDC_SetMod.sStr(C_rec.C095) + MDC_SetMod.sStr(C_rec.C096) + MDC_SetMod.sStr(C_rec.C097) + MDC_SetMod.sStr(C_rec.C098) + MDC_SetMod.sStr(C_rec.C099) + MDC_SetMod.sStr(C_rec.C100) + MDC_SetMod.sStr(C_rec.C101) + MDC_SetMod.sStr(C_rec.C102) + MDC_SetMod.sStr(C_rec.C103) + MDC_SetMod.sStr(C_rec.C104) + MDC_SetMod.sStr(C_rec.C105) + MDC_SetMod.sStr(C_rec.C106) + MDC_SetMod.sStr(C_rec.C108) + MDC_SetMod.sStr(C_rec.C109) + MDC_SetMod.sStr(C_rec.C110) + MDC_SetMod.sStr(C_rec.C111) + MDC_SetMod.sStr(C_rec.C112) + MDC_SetMod.sStr(C_rec.C113) + MDC_SetMod.sStr(C_rec.C114) + MDC_SetMod.sStr(C_rec.C115) + MDC_SetMod.sStr(C_rec.C116) + MDC_SetMod.sStr(C_rec.C118) + MDC_SetMod.sStr(C_rec.C119) + MDC_SetMod.sStr(C_rec.C120A) + MDC_SetMod.sStr(C_rec.C120B) + MDC_SetMod.sStr(C_rec.C121A) + MDC_SetMod.sStr(C_rec.C121B) + MDC_SetMod.sStr(C_rec.C122A) + MDC_SetMod.sStr(C_rec.C122B) + MDC_SetMod.sStr(C_rec.C123A) + MDC_SetMod.sStr(C_rec.C123B) + MDC_SetMod.sStr(C_rec.C124A) + MDC_SetMod.sStr(C_rec.C124B) + MDC_SetMod.sStr(C_rec.C125A) + MDC_SetMod.sStr(C_rec.C125B) + MDC_SetMod.sStr(C_rec.C126A) + MDC_SetMod.sStr(C_rec.C126B) + MDC_SetMod.sStr(C_rec.C127A) + MDC_SetMod.sStr(C_rec.C127B) + MDC_SetMod.sStr(C_rec.C128A) + MDC_SetMod.sStr(C_rec.C128B) + MDC_SetMod.sStr(C_rec.C129A) + MDC_SetMod.sStr(C_rec.C129B) + MDC_SetMod.sStr(C_rec.C130A) + MDC_SetMod.sStr(C_rec.C130B) + MDC_SetMod.sStr(C_rec.C131A) + MDC_SetMod.sStr(C_rec.C131B) + MDC_SetMod.sStr(C_rec.C132A) + MDC_SetMod.sStr(C_rec.C132B) + MDC_SetMod.sStr(C_rec.C133A) + MDC_SetMod.sStr(C_rec.C133B) + MDC_SetMod.sStr(C_rec.C134A) + MDC_SetMod.sStr(C_rec.C134B) + MDC_SetMod.sStr(C_rec.C135) + MDC_SetMod.sStr(C_rec.C137) + MDC_SetMod.sStr(C_rec.C138) + MDC_SetMod.sStr(C_rec.C139) + MDC_SetMod.sStr(C_rec.C140) + MDC_SetMod.sStr(C_rec.C141) + MDC_SetMod.sStr(C_rec.C142A) + MDC_SetMod.sStr(C_rec.C142B) + MDC_SetMod.sStr(C_rec.C143) + MDC_SetMod.sStr(C_rec.C145) + MDC_SetMod.sStr(C_rec.C146A) + MDC_SetMod.sStr(C_rec.C146B) + MDC_SetMod.sStr(C_rec.C146C) + MDC_SetMod.sStr(C_rec.C147A) + MDC_SetMod.sStr(C_rec.C147B) + MDC_SetMod.sStr(C_rec.C147C) + MDC_SetMod.sStr(C_rec.C148A) + MDC_SetMod.sStr(C_rec.C148B) + MDC_SetMod.sStr(C_rec.C148C) + MDC_SetMod.sStr(C_rec.C149A_1) + MDC_SetMod.sStr(C_rec.C149A_2) + MDC_SetMod.sStr(C_rec.C149B_1) + MDC_SetMod.sStr(C_rec.C149B_2) + MDC_SetMod.sStr(C_rec.C149C_1) + MDC_SetMod.sStr(C_rec.C149C_2) + MDC_SetMod.sStr(C_rec.C150));


//            // / D레코드: 종전근무처 레코드
//            if (Conversion.Val(C_rec.C006) > 0)
//            {
//                if (File_Create_D_record == false)
//                {
//                    ErrNum = 2;
//                    goto Error_Message;
//                }
//            }

//            // / E레코드: 부양가족 레코드
//            if (File_Create_E_record == false)
//            {
//                ErrNum = 3;
//                goto Error_Message;
//            }

//            // / F레코드
//            if (File_Create_F_record == false)
//            {
//                ErrNum = 4;
//                goto Error_Message;
//            }

//            // / G레코드
//            if (File_Create_G_record == false)
//            {
//                ErrNum = 5;
//                goto Error_Message;
//            }

//            // / H레코드
//            if (File_Create_H_record == false)
//            {
//                ErrNum = 6;
//                goto Error_Message;
//            }

//            // / I레코드
//            if (File_Create_I_record == false)
//            {
//                ErrNum = 7;
//                goto Error_Message;
//            }


//            ProgressBar01.VALUE = ProgressBar01.VALUE + 1;
//            ProgressBar01.Text = ProgressBar01.VALUE + "/" + oRecordSet.RecordCount + "건 작성중........!";


//            oRecordSet.MoveNext();
//        }

//        ProgressBar01.Stop();
//        // UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        ProgressBar01 = null/* TODO Change to default(_) if this is not a reference type */;

//        if (System.Convert.ToBoolean(CheckC) == false)
//            File_Create_C_record = true;
//        else
//            File_Create_C_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("C레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("D레코드(종전근무처 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("E레코드(부양가족 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 4)
//            Sbo_Application.StatusBar.SetText("F레코드(연금.저축 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 5)
//            Sbo_Application.StatusBar.SetText("G레코드(월세액.주택자료 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 6)
//            Sbo_Application.StatusBar.SetText("H레코드(기부금조정명세 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 7)
//            Sbo_Application.StatusBar.SetText("I레코드(해당연도 기부금명세 레코드) 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("C레코드오류: " + Information.Err.Description, ref false);
//        File_Create_C_record = false;
//    }

//    private bool File_Create_D_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 224291
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

//        // / D_RECORE QUERY
//        sQry = "EXEC PH_PY980_D '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";

//        oRecordSet.DoQuery(sQry);
//        if (oRecordSet.RecordCount == 0)
//        {
//            ErrNum = 1;
//            goto Error_Message;
//        }

//        while (!oRecordSet.EOF)
//        {

//            // D RECORD MOVE

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D001 = oRecordSet.Fields.Item("D001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D002 = oRecordSet.Fields.Item("D002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D003 = oRecordSet.Fields.Item("D003").VALUE;
//            D_rec.D004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(D_rec.D004))); // / C레코드의 일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D005 = oRecordSet.Fields.Item("D005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D006 = oRecordSet.Fields.Item("D006").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D007 = oRecordSet.Fields.Item("D007").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D008 = oRecordSet.Fields.Item("D008").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D009 = oRecordSet.Fields.Item("D009").VALUE;
//            D_rec.D010 = VB6.Format(oRecordSet.Fields.Item("D010").VALUE, new string("0", Strings.Len(D_rec.D010)));
//            D_rec.D011 = VB6.Format(oRecordSet.Fields.Item("D011").VALUE, new string("0", Strings.Len(D_rec.D011)));
//            D_rec.D012 = VB6.Format(oRecordSet.Fields.Item("D012").VALUE, new string("0", Strings.Len(D_rec.D012)));
//            D_rec.D013 = VB6.Format(oRecordSet.Fields.Item("D013").VALUE, new string("0", Strings.Len(D_rec.D013)));
//            D_rec.D014 = VB6.Format(oRecordSet.Fields.Item("D014").VALUE, new string("0", Strings.Len(D_rec.D014)));
//            D_rec.D015 = VB6.Format(oRecordSet.Fields.Item("D015").VALUE, new string("0", Strings.Len(D_rec.D015)));
//            D_rec.D016 = VB6.Format(oRecordSet.Fields.Item("D016").VALUE, new string("0", Strings.Len(D_rec.D016)));
//            D_rec.D017 = VB6.Format(oRecordSet.Fields.Item("D017").VALUE, new string("0", Strings.Len(D_rec.D017)));
//            D_rec.D018 = VB6.Format(oRecordSet.Fields.Item("D018").VALUE, new string("0", Strings.Len(D_rec.D018)));
//            D_rec.D019 = VB6.Format(oRecordSet.Fields.Item("D019").VALUE, new string("0", Strings.Len(D_rec.D019)));
//            D_rec.D020 = VB6.Format(oRecordSet.Fields.Item("D020").VALUE, new string("0", Strings.Len(D_rec.D020)));
//            D_rec.D021 = VB6.Format(oRecordSet.Fields.Item("D021").VALUE, new string("0", Strings.Len(D_rec.D021)));

//            D_rec.D023 = VB6.Format(oRecordSet.Fields.Item("D023").VALUE, new string("0", Strings.Len(D_rec.D023)));
//            D_rec.D024 = VB6.Format(oRecordSet.Fields.Item("D024").VALUE, new string("0", Strings.Len(D_rec.D024)));
//            D_rec.D025 = VB6.Format(oRecordSet.Fields.Item("D025").VALUE, new string("0", Strings.Len(D_rec.D025)));
//            D_rec.D026 = VB6.Format(oRecordSet.Fields.Item("D026").VALUE, new string("0", Strings.Len(D_rec.D026)));
//            D_rec.D027 = VB6.Format(oRecordSet.Fields.Item("D027").VALUE, new string("0", Strings.Len(D_rec.D027)));
//            D_rec.D028 = VB6.Format(oRecordSet.Fields.Item("D028").VALUE, new string("0", Strings.Len(D_rec.D028)));
//            D_rec.D029 = VB6.Format(oRecordSet.Fields.Item("D029").VALUE, new string("0", Strings.Len(D_rec.D029)));
//            D_rec.D030 = VB6.Format(oRecordSet.Fields.Item("D030").VALUE, new string("0", Strings.Len(D_rec.D030)));
//            D_rec.D031 = VB6.Format(oRecordSet.Fields.Item("D031").VALUE, new string("0", Strings.Len(D_rec.D031)));
//            D_rec.D032 = VB6.Format(oRecordSet.Fields.Item("D032").VALUE, new string("0", Strings.Len(D_rec.D032)));
//            D_rec.D033 = VB6.Format(oRecordSet.Fields.Item("D033").VALUE, new string("0", Strings.Len(D_rec.D033)));
//            D_rec.D034 = VB6.Format(oRecordSet.Fields.Item("D034").VALUE, new string("0", Strings.Len(D_rec.D034)));
//            D_rec.D035 = VB6.Format(oRecordSet.Fields.Item("D035").VALUE, new string("0", Strings.Len(D_rec.D035)));
//            D_rec.D036 = VB6.Format(oRecordSet.Fields.Item("D036").VALUE, new string("0", Strings.Len(D_rec.D036)));
//            D_rec.D037 = VB6.Format(oRecordSet.Fields.Item("D037").VALUE, new string("0", Strings.Len(D_rec.D037)));
//            D_rec.D038 = VB6.Format(oRecordSet.Fields.Item("D038").VALUE, new string("0", Strings.Len(D_rec.D038)));
//            D_rec.D039 = VB6.Format(oRecordSet.Fields.Item("D039").VALUE, new string("0", Strings.Len(D_rec.D039)));
//            D_rec.D040 = VB6.Format(oRecordSet.Fields.Item("D040").VALUE, new string("0", Strings.Len(D_rec.D040)));
//            D_rec.D041 = VB6.Format(oRecordSet.Fields.Item("D041").VALUE, new string("0", Strings.Len(D_rec.D041)));
//            D_rec.D042 = VB6.Format(oRecordSet.Fields.Item("D042").VALUE, new string("0", Strings.Len(D_rec.D042)));
//            D_rec.D043 = VB6.Format(oRecordSet.Fields.Item("D043").VALUE, new string("0", Strings.Len(D_rec.D043)));
//            D_rec.D044 = VB6.Format(oRecordSet.Fields.Item("D044").VALUE, new string("0", Strings.Len(D_rec.D044)));
//            D_rec.D045 = VB6.Format(oRecordSet.Fields.Item("D045").VALUE, new string("0", Strings.Len(D_rec.D045)));
//            D_rec.D046 = VB6.Format(oRecordSet.Fields.Item("D046").VALUE, new string("0", Strings.Len(D_rec.D046)));
//            D_rec.D047 = VB6.Format(oRecordSet.Fields.Item("D047").VALUE, new string("0", Strings.Len(D_rec.D047)));
//            D_rec.D048 = VB6.Format(oRecordSet.Fields.Item("D048").VALUE, new string("0", Strings.Len(D_rec.D048)));
//            D_rec.D049 = VB6.Format(oRecordSet.Fields.Item("D049").VALUE, new string("0", Strings.Len(D_rec.D049)));
//            D_rec.D050A = VB6.Format(oRecordSet.Fields.Item("D050A").VALUE, new string("0", Strings.Len(D_rec.D050A)));
//            D_rec.D050B = VB6.Format(oRecordSet.Fields.Item("D050B").VALUE, new string("0", Strings.Len(D_rec.D050B)));
//            D_rec.D050C = VB6.Format(oRecordSet.Fields.Item("D050C").VALUE, new string("0", Strings.Len(D_rec.D050C)));
//            D_rec.D051 = VB6.Format(oRecordSet.Fields.Item("D051").VALUE, new string("0", Strings.Len(D_rec.D051)));
//            D_rec.D052 = VB6.Format(oRecordSet.Fields.Item("D052").VALUE, new string("0", Strings.Len(D_rec.D052)));
//            D_rec.D053A = VB6.Format(oRecordSet.Fields.Item("D053A").VALUE, new string("0", Strings.Len(D_rec.D053A)));
//            D_rec.D053B = VB6.Format(oRecordSet.Fields.Item("D053B").VALUE, new string("0", Strings.Len(D_rec.D053B)));
//            D_rec.D053C = VB6.Format(oRecordSet.Fields.Item("D053C").VALUE, new string("0", Strings.Len(D_rec.D053C)));
//            D_rec.D053D = VB6.Format(oRecordSet.Fields.Item("D053D").VALUE, new string("0", Strings.Len(D_rec.D053D)));
//            D_rec.D054 = VB6.Format(oRecordSet.Fields.Item("D054").VALUE, new string("0", Strings.Len(D_rec.D054)));
//            D_rec.D055 = VB6.Format(oRecordSet.Fields.Item("D055").VALUE, new string("0", Strings.Len(D_rec.D055)));

//            D_rec.D057 = VB6.Format(oRecordSet.Fields.Item("D057").VALUE, new string("0", Strings.Len(D_rec.D057)));
//            D_rec.D058 = VB6.Format(oRecordSet.Fields.Item("D058").VALUE, new string("0", Strings.Len(D_rec.D058)));
//            D_rec.D059A = VB6.Format(oRecordSet.Fields.Item("D059A").VALUE, new string("0", Strings.Len(D_rec.D059A)));
//            D_rec.D059B = VB6.Format(oRecordSet.Fields.Item("D059B").VALUE, new string("0", Strings.Len(D_rec.D059B)));
//            D_rec.D059C = VB6.Format(oRecordSet.Fields.Item("D059C").VALUE, new string("0", Strings.Len(D_rec.D059C)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D060 = oRecordSet.Fields.Item("D060").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            D_rec.D061 = oRecordSet.Fields.Item("D061").VALUE;

//            PrintLine(1, MDC_SetMod.sStr(D_rec.D001) + MDC_SetMod.sStr(D_rec.D002) + MDC_SetMod.sStr(D_rec.D003) + MDC_SetMod.sStr(D_rec.D004) + MDC_SetMod.sStr(D_rec.D005) + MDC_SetMod.sStr(D_rec.D006) + MDC_SetMod.sStr(D_rec.D007) + MDC_SetMod.sStr(D_rec.D008) + MDC_SetMod.sStr(D_rec.D009) + MDC_SetMod.sStr(D_rec.D010) + MDC_SetMod.sStr(D_rec.D011) + MDC_SetMod.sStr(D_rec.D012) + MDC_SetMod.sStr(D_rec.D013) + MDC_SetMod.sStr(D_rec.D014) + MDC_SetMod.sStr(D_rec.D015) + MDC_SetMod.sStr(D_rec.D016) + MDC_SetMod.sStr(D_rec.D017) + MDC_SetMod.sStr(D_rec.D018) + MDC_SetMod.sStr(D_rec.D019) + MDC_SetMod.sStr(D_rec.D020) + MDC_SetMod.sStr(D_rec.D021) + MDC_SetMod.sStr(D_rec.D023) + MDC_SetMod.sStr(D_rec.D024) + MDC_SetMod.sStr(D_rec.D025) + MDC_SetMod.sStr(D_rec.D026) + MDC_SetMod.sStr(D_rec.D027) + MDC_SetMod.sStr(D_rec.D028) + MDC_SetMod.sStr(D_rec.D029) + MDC_SetMod.sStr(D_rec.D030) + MDC_SetMod.sStr(D_rec.D031) + MDC_SetMod.sStr(D_rec.D032) + MDC_SetMod.sStr(D_rec.D033) + MDC_SetMod.sStr(D_rec.D034) + MDC_SetMod.sStr(D_rec.D035) + MDC_SetMod.sStr(D_rec.D036) + MDC_SetMod.sStr(D_rec.D037) + MDC_SetMod.sStr(D_rec.D038) + MDC_SetMod.sStr(D_rec.D039) + MDC_SetMod.sStr(D_rec.D040) + MDC_SetMod.sStr(D_rec.D041) + MDC_SetMod.sStr(D_rec.D042) + MDC_SetMod.sStr(D_rec.D043) + MDC_SetMod.sStr(D_rec.D044) + MDC_SetMod.sStr(D_rec.D045) + MDC_SetMod.sStr(D_rec.D046) + MDC_SetMod.sStr(D_rec.D047) + MDC_SetMod.sStr(D_rec.D048) + MDC_SetMod.sStr(D_rec.D049) + MDC_SetMod.sStr(D_rec.D050A) + MDC_SetMod.sStr(D_rec.D050B) + MDC_SetMod.sStr(D_rec.D050C) + MDC_SetMod.sStr(D_rec.D051) + MDC_SetMod.sStr(D_rec.D052) + MDC_SetMod.sStr(D_rec.D053A) + MDC_SetMod.sStr(D_rec.D053B) + MDC_SetMod.sStr(D_rec.D053C) + MDC_SetMod.sStr(D_rec.D053D) + MDC_SetMod.sStr(D_rec.D054) + MDC_SetMod.sStr(D_rec.D055) + MDC_SetMod.sStr(D_rec.D057) + MDC_SetMod.sStr(D_rec.D058) + MDC_SetMod.sStr(D_rec.D059A) + MDC_SetMod.sStr(D_rec.D059B) + MDC_SetMod.sStr(D_rec.D059C) + MDC_SetMod.sStr(D_rec.D060) + MDC_SetMod.sStr(D_rec.D061));

//            oRecordSet.MoveNext();
//        }

//        if (System.Convert.ToBoolean(CheckD) == false)
//            File_Create_D_record = true;
//        else
//            File_Create_D_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("종전근무지D레코드가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Matrix_AddRow("D레코드오류: " + Information.Err.Description, ref false);
//        File_Create_D_record = false;
//    }

//    private bool File_Create_E_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 235997
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

//        // / E_RECORE QUERY
//        sQry = "EXEC PH_PY980_E '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";
//        oRecordSet.DoQuery(sQry);

//        if (oRecordSet.RecordCount > 0)
//        {
//            BUYCNT = 0; // 가족수
//            FAMCNT = 1; // E레코드일련번호

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            E_rec.E001 = oRecordSet.Fields.Item("E001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            E_rec.E002 = oRecordSet.Fields.Item("E002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            E_rec.E003 = oRecordSet.Fields.Item("E003").VALUE;
//            E_rec.E004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(E_rec.E004))); // / C레코드의 일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            E_rec.E005 = oRecordSet.Fields.Item("E005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            E_rec.E006 = oRecordSet.Fields.Item("E006").VALUE;

//            // TOTBOH = 0: TOTMED = 0: TOTEDC = 0: TOTCAD = 0: TOTCSH = 0: TOTGBU = 0

//            while (!oRecordSet.EOF)
//            {
//                BUYCNT = BUYCNT + 1; // / 해당사원의 부양가족일련번호
//                                     // /초기화
//                if (BUYCNT == 1)
//                {
//                    for (i = 1; i <= 5; i++)
//                    {
//                        E_rec.E007[i] = Strings.Space(Strings.Len(E_rec.E007[i]));
//                        E_rec.E008[i] = Strings.Space(Strings.Len(E_rec.E008[i]));
//                        E_rec.E009[i] = Strings.Space(Strings.Len(E_rec.E009[i]));
//                        E_rec.E010[i] = Strings.Space(Strings.Len(E_rec.E010[i]));
//                        E_rec.E011[i] = Strings.Space(Strings.Len(E_rec.E011[i]));
//                        E_rec.E012[i] = Strings.Space(Strings.Len(E_rec.E012[i]));
//                        E_rec.E013[i] = Strings.Space(Strings.Len(E_rec.E013[i]));
//                        E_rec.E014[i] = Strings.Space(Strings.Len(E_rec.E014[i]));
//                        E_rec.E015[i] = Strings.Space(Strings.Len(E_rec.E015[i]));
//                        E_rec.E016[i] = Strings.Space(Strings.Len(E_rec.E016[i]));
//                        E_rec.E017[i] = Strings.Space(Strings.Len(E_rec.E017[i]));
//                        E_rec.E018[i] = Strings.Space(Strings.Len(E_rec.E018[i]));

//                        E_rec.E019[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E019[i])));
//                        E_rec.E020[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E020[i])));
//                        E_rec.E021[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E021[i])));
//                        E_rec.E022[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E022[i])));
//                        E_rec.E023[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E023[i])));
//                        E_rec.E024[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E024[i])));
//                        E_rec.E025[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E025[i])));
//                        E_rec.E026[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E026[i])));
//                        E_rec.E027[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E027[i])));
//                        E_rec.E028[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E028[i])));
//                        E_rec.E029[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E029[i])));
//                        E_rec.E030[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E030[i])));
//                        E_rec.E031[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E031[i])));
//                        E_rec.E032[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E032[i])));
//                        E_rec.E033[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E033[i])));
//                        E_rec.E034[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E034[i])));
//                        E_rec.E035[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E035[i])));
//                        E_rec.E036[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E036[i])));
//                        E_rec.E037[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E037[i])));
//                        E_rec.E038[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E038[i])));
//                        E_rec.E039[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E039[i])));
//                        E_rec.E040[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E040[i])));
//                        E_rec.E041[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E041[i])));
//                        E_rec.E042[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E042[i])));
//                        E_rec.E043[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E043[i])));
//                        E_rec.E044[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E044[i])));
//                        E_rec.E045[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E045[i])));
//                        E_rec.E046[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E046[i])));
//                        E_rec.E047[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E047[i])));
//                        E_rec.E048[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E048[i])));
//                        E_rec.E049[i] = VB6.Format(0, new string("0", Strings.Len(E_rec.E049[i])));
//                    }
//                }

//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E007[BUYCNT] = oRecordSet.Fields.Item("E007").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E008[BUYCNT] = oRecordSet.Fields.Item("E008").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E009[BUYCNT] = oRecordSet.Fields.Item("E009").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E010[BUYCNT] = oRecordSet.Fields.Item("E010").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E011[BUYCNT] = oRecordSet.Fields.Item("E011").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E012[BUYCNT] = oRecordSet.Fields.Item("E012").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E013[BUYCNT] = oRecordSet.Fields.Item("E013").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E014[BUYCNT] = oRecordSet.Fields.Item("E014").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E015[BUYCNT] = oRecordSet.Fields.Item("E015").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E016[BUYCNT] = oRecordSet.Fields.Item("E016").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E017[BUYCNT] = oRecordSet.Fields.Item("E017").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                E_rec.E018[BUYCNT] = oRecordSet.Fields.Item("E018").VALUE;
//                E_rec.E019[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E019").VALUE, new string("0", Strings.Len(E_rec.E019[BUYCNT])));
//                E_rec.E020[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E020").VALUE, new string("0", Strings.Len(E_rec.E020[BUYCNT])));
//                E_rec.E021[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E021").VALUE, new string("0", Strings.Len(E_rec.E021[BUYCNT])));
//                E_rec.E022[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E022").VALUE, new string("0", Strings.Len(E_rec.E022[BUYCNT])));
//                E_rec.E023[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E023").VALUE, new string("0", Strings.Len(E_rec.E023[BUYCNT])));
//                E_rec.E024[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E024").VALUE, new string("0", Strings.Len(E_rec.E024[BUYCNT])));
//                E_rec.E025[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E025").VALUE, new string("0", Strings.Len(E_rec.E025[BUYCNT])));
//                E_rec.E026[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E026").VALUE, new string("0", Strings.Len(E_rec.E026[BUYCNT])));
//                E_rec.E027[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E027").VALUE, new string("0", Strings.Len(E_rec.E027[BUYCNT])));
//                E_rec.E028[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E028").VALUE, new string("0", Strings.Len(E_rec.E028[BUYCNT])));
//                E_rec.E029[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E029").VALUE, new string("0", Strings.Len(E_rec.E029[BUYCNT])));
//                E_rec.E030[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E030").VALUE, new string("0", Strings.Len(E_rec.E030[BUYCNT])));
//                E_rec.E031[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E031").VALUE, new string("0", Strings.Len(E_rec.E031[BUYCNT])));
//                E_rec.E032[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E032").VALUE, new string("0", Strings.Len(E_rec.E032[BUYCNT])));
//                E_rec.E033[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E033").VALUE, new string("0", Strings.Len(E_rec.E033[BUYCNT])));
//                E_rec.E034[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E034").VALUE, new string("0", Strings.Len(E_rec.E034[BUYCNT])));
//                E_rec.E035[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E035").VALUE, new string("0", Strings.Len(E_rec.E035[BUYCNT])));
//                E_rec.E036[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E036").VALUE, new string("0", Strings.Len(E_rec.E036[BUYCNT])));
//                E_rec.E037[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E037").VALUE, new string("0", Strings.Len(E_rec.E037[BUYCNT])));
//                E_rec.E038[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E038").VALUE, new string("0", Strings.Len(E_rec.E038[BUYCNT])));
//                E_rec.E039[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E039").VALUE, new string("0", Strings.Len(E_rec.E039[BUYCNT])));
//                E_rec.E040[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E040").VALUE, new string("0", Strings.Len(E_rec.E040[BUYCNT])));
//                E_rec.E041[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E041").VALUE, new string("0", Strings.Len(E_rec.E041[BUYCNT])));
//                E_rec.E042[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E042").VALUE, new string("0", Strings.Len(E_rec.E042[BUYCNT])));
//                E_rec.E043[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E043").VALUE, new string("0", Strings.Len(E_rec.E043[BUYCNT])));
//                E_rec.E044[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E044").VALUE, new string("0", Strings.Len(E_rec.E044[BUYCNT])));
//                E_rec.E045[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E045").VALUE, new string("0", Strings.Len(E_rec.E045[BUYCNT])));
//                E_rec.E046[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E046").VALUE, new string("0", Strings.Len(E_rec.E046[BUYCNT])));
//                E_rec.E047[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E047").VALUE, new string("0", Strings.Len(E_rec.E047[BUYCNT])));
//                E_rec.E048[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E048").VALUE, new string("0", Strings.Len(E_rec.E048[BUYCNT])));
//                E_rec.E049[BUYCNT] = VB6.Format(oRecordSet.Fields.Item("E049").VALUE, new string("0", Strings.Len(E_rec.E049[BUYCNT])));


//                oRecordSet.MoveNext();

//                // If BUYCNT = 5 Then    '5개면 인쇄
//                if (BUYCNT == 5 | oRecordSet.EOF)
//                {
//                    E_rec.E222 = VB6.Format(FAMCNT, new string("0", Strings.Len(E_rec.E222))); // --일련번호
//                                                                                               // / E레코드삽입
//                    PrintLine(1, MDC_SetMod.sStr(E_rec.E001) + MDC_SetMod.sStr(E_rec.E002) + MDC_SetMod.sStr(E_rec.E003) + MDC_SetMod.sStr(E_rec.E004) + MDC_SetMod.sStr(E_rec.E005) + MDC_SetMod.sStr(E_rec.E006) + MDC_SetMod.sStr(E_rec.E007[1]) + MDC_SetMod.sStr(E_rec.E008[1]) + MDC_SetMod.sStr(E_rec.E009[1]) + MDC_SetMod.sStr(E_rec.E010[1]) + MDC_SetMod.sStr(E_rec.E011[1]) + MDC_SetMod.sStr(E_rec.E012[1]) + MDC_SetMod.sStr(E_rec.E013[1]) + MDC_SetMod.sStr(E_rec.E014[1]) + MDC_SetMod.sStr(E_rec.E015[1]) + MDC_SetMod.sStr(E_rec.E016[1]) + MDC_SetMod.sStr(E_rec.E017[1]) + MDC_SetMod.sStr(E_rec.E018[1]) + MDC_SetMod.sStr(E_rec.E019[1]) + MDC_SetMod.sStr(E_rec.E020[1]) + MDC_SetMod.sStr(E_rec.E021[1]) + MDC_SetMod.sStr(E_rec.E022[1]) + MDC_SetMod.sStr(E_rec.E023[1]) + MDC_SetMod.sStr(E_rec.E024[1]) + MDC_SetMod.sStr(E_rec.E025[1]) + MDC_SetMod.sStr(E_rec.E026[1]) + MDC_SetMod.sStr(E_rec.E027[1]) + MDC_SetMod.sStr(E_rec.E028[1]) + MDC_SetMod.sStr(E_rec.E029[1]) + MDC_SetMod.sStr(E_rec.E030[1]) + MDC_SetMod.sStr(E_rec.E031[1]) + MDC_SetMod.sStr(E_rec.E032[1]) + MDC_SetMod.sStr(E_rec.E033[1]) + MDC_SetMod.sStr(E_rec.E034[1]) + MDC_SetMod.sStr(E_rec.E035[1]) + MDC_SetMod.sStr(E_rec.E036[1]) + MDC_SetMod.sStr(E_rec.E037[1]) + MDC_SetMod.sStr(E_rec.E038[1]) + MDC_SetMod.sStr(E_rec.E039[1]) + MDC_SetMod.sStr(E_rec.E040[1]) + MDC_SetMod.sStr(E_rec.E041[1]) + MDC_SetMod.sStr(E_rec.E042[1]) + MDC_SetMod.sStr(E_rec.E043[1]) + MDC_SetMod.sStr(E_rec.E044[1]) + MDC_SetMod.sStr(E_rec.E045[1]) + MDC_SetMod.sStr(E_rec.E046[1]) + MDC_SetMod.sStr(E_rec.E047[1]) + MDC_SetMod.sStr(E_rec.E048[1]) + MDC_SetMod.sStr(E_rec.E049[1]) + MDC_SetMod.sStr(E_rec.E007[2]) + MDC_SetMod.sStr(E_rec.E008[2]) + MDC_SetMod.sStr(E_rec.E009[2]) + MDC_SetMod.sStr(E_rec.E010[2]) + MDC_SetMod.sStr(E_rec.E011[2]) + MDC_SetMod.sStr(E_rec.E012[2]) + MDC_SetMod.sStr(E_rec.E013[2]) + MDC_SetMod.sStr(E_rec.E014[2]) + MDC_SetMod.sStr(E_rec.E015[2]) + MDC_SetMod.sStr(E_rec.E016[2]) + MDC_SetMod.sStr(E_rec.E017[2]) + MDC_SetMod.sStr(E_rec.E018[2]) + MDC_SetMod.sStr(E_rec.E019[2]) + MDC_SetMod.sStr(E_rec.E020[2]) + MDC_SetMod.sStr(E_rec.E021[2]) + MDC_SetMod.sStr(E_rec.E022[2]) + MDC_SetMod.sStr(E_rec.E023[2]) + MDC_SetMod.sStr(E_rec.E024[2]) + MDC_SetMod.sStr(E_rec.E025[2]) + MDC_SetMod.sStr(E_rec.E026[2]) + MDC_SetMod.sStr(E_rec.E027[2]) + MDC_SetMod.sStr(E_rec.E028[2]) + MDC_SetMod.sStr(E_rec.E029[2]) + MDC_SetMod.sStr(E_rec.E030[2]) + MDC_SetMod.sStr(E_rec.E031[2]) + MDC_SetMod.sStr(E_rec.E032[2]) + MDC_SetMod.sStr(E_rec.E033[2]) + MDC_SetMod.sStr(E_rec.E034[2]) + MDC_SetMod.sStr(E_rec.E035[2]) + MDC_SetMod.sStr(E_rec.E036[2]) + MDC_SetMod.sStr(E_rec.E037[2]) + MDC_SetMod.sStr(E_rec.E038[2]) + MDC_SetMod.sStr(E_rec.E039[2]) + MDC_SetMod.sStr(E_rec.E040[2]) + MDC_SetMod.sStr(E_rec.E041[2]) + MDC_SetMod.sStr(E_rec.E042[2]) + MDC_SetMod.sStr(E_rec.E043[2]) + MDC_SetMod.sStr(E_rec.E044[2]) + MDC_SetMod.sStr(E_rec.E045[2]) + MDC_SetMod.sStr(E_rec.E046[2]) + MDC_SetMod.sStr(E_rec.E047[2]) + MDC_SetMod.sStr(E_rec.E048[2]) + MDC_SetMod.sStr(E_rec.E049[2]) + MDC_SetMod.sStr(E_rec.E007[3]) + MDC_SetMod.sStr(E_rec.E008[3]) + MDC_SetMod.sStr(E_rec.E009[3]) + MDC_SetMod.sStr(E_rec.E010[3]) + MDC_SetMod.sStr(E_rec.E011[3]) + MDC_SetMod.sStr(E_rec.E012[3]) + MDC_SetMod.sStr(E_rec.E013[3]) + MDC_SetMod.sStr(E_rec.E014[3]) + MDC_SetMod.sStr(E_rec.E015[3]) + MDC_SetMod.sStr(E_rec.E016[3]) + MDC_SetMod.sStr(E_rec.E017[3]) + MDC_SetMod.sStr(E_rec.E018[3]) + MDC_SetMod.sStr(E_rec.E019[3]) + MDC_SetMod.sStr(E_rec.E020[3]) + MDC_SetMod.sStr(E_rec.E021[3]) + MDC_SetMod.sStr(E_rec.E022[3]) + MDC_SetMod.sStr(E_rec.E023[3]) + MDC_SetMod.sStr(E_rec.E024[3]) + MDC_SetMod.sStr(E_rec.E025[3]) + MDC_SetMod.sStr(E_rec.E026[3]) + MDC_SetMod.sStr(E_rec.E027[3]) + MDC_SetMod.sStr(E_rec.E028[3]) + MDC_SetMod.sStr(E_rec.E029[3]) + MDC_SetMod.sStr(E_rec.E030[3]) + MDC_SetMod.sStr(E_rec.E031[3]) + MDC_SetMod.sStr(E_rec.E032[3]) + MDC_SetMod.sStr(E_rec.E033[3]) + MDC_SetMod.sStr(E_rec.E034[3]) + MDC_SetMod.sStr(E_rec.E035[3]) + MDC_SetMod.sStr(E_rec.E036[3]) + MDC_SetMod.sStr(E_rec.E037[3]) + MDC_SetMod.sStr(E_rec.E038[3]) + MDC_SetMod.sStr(E_rec.E039[3]) + MDC_SetMod.sStr(E_rec.E040[3]) + MDC_SetMod.sStr(E_rec.E041[3]) + MDC_SetMod.sStr(E_rec.E042[3]) + MDC_SetMod.sStr(E_rec.E043[3]) + MDC_SetMod.sStr(E_rec.E044[3]) + MDC_SetMod.sStr(E_rec.E045[3]) + MDC_SetMod.sStr(E_rec.E046[3]) + MDC_SetMod.sStr(E_rec.E047[3]) + MDC_SetMod.sStr(E_rec.E048[3]) + MDC_SetMod.sStr(E_rec.E049[3]) + MDC_SetMod.sStr(E_rec.E007[4]) + MDC_SetMod.sStr(E_rec.E008[4]) + MDC_SetMod.sStr(E_rec.E009[4]) + MDC_SetMod.sStr(E_rec.E010[4]) + MDC_SetMod.sStr(E_rec.E011[4]) + MDC_SetMod.sStr(E_rec.E012[4]) + MDC_SetMod.sStr(E_rec.E013[4]) + MDC_SetMod.sStr(E_rec.E014[4]) + MDC_SetMod.sStr(E_rec.E015[4]) + MDC_SetMod.sStr(E_rec.E016[4]) + MDC_SetMod.sStr(E_rec.E017[4]) + MDC_SetMod.sStr(E_rec.E018[4]) + MDC_SetMod.sStr(E_rec.E019[4]) + MDC_SetMod.sStr(E_rec.E020[4]) + MDC_SetMod.sStr(E_rec.E021[4]) + MDC_SetMod.sStr(E_rec.E022[4]) + MDC_SetMod.sStr(E_rec.E023[4]) + MDC_SetMod.sStr(E_rec.E024[4]) + MDC_SetMod.sStr(E_rec.E025[4]) + MDC_SetMod.sStr(E_rec.E026[4]) + MDC_SetMod.sStr(E_rec.E027[4]) + MDC_SetMod.sStr(E_rec.E028[4]) + MDC_SetMod.sStr(E_rec.E029[4]) + MDC_SetMod.sStr(E_rec.E030[4]) + MDC_SetMod.sStr(E_rec.E031[4]) + MDC_SetMod.sStr(E_rec.E032[4]) + MDC_SetMod.sStr(E_rec.E033[4]) + MDC_SetMod.sStr(E_rec.E034[4]) + MDC_SetMod.sStr(E_rec.E035[4]) + MDC_SetMod.sStr(E_rec.E036[4]) + MDC_SetMod.sStr(E_rec.E037[4]) + MDC_SetMod.sStr(E_rec.E038[4]) + MDC_SetMod.sStr(E_rec.E039[4]) + MDC_SetMod.sStr(E_rec.E040[4]) + MDC_SetMod.sStr(E_rec.E041[4]) + MDC_SetMod.sStr(E_rec.E042[4]) + MDC_SetMod.sStr(E_rec.E043[4]) + MDC_SetMod.sStr(E_rec.E044[4]) + MDC_SetMod.sStr(E_rec.E045[4]) + MDC_SetMod.sStr(E_rec.E046[4]) + MDC_SetMod.sStr(E_rec.E047[4]) + MDC_SetMod.sStr(E_rec.E048[4]) + MDC_SetMod.sStr(E_rec.E049[4]) + MDC_SetMod.sStr(E_rec.E007[5]) + MDC_SetMod.sStr(E_rec.E008[5]) + MDC_SetMod.sStr(E_rec.E009[5]) + MDC_SetMod.sStr(E_rec.E010[5]) + MDC_SetMod.sStr(E_rec.E011[5]) + MDC_SetMod.sStr(E_rec.E012[5]) + MDC_SetMod.sStr(E_rec.E013[5]) + MDC_SetMod.sStr(E_rec.E014[5]) + MDC_SetMod.sStr(E_rec.E015[5]) + MDC_SetMod.sStr(E_rec.E016[5]) + MDC_SetMod.sStr(E_rec.E017[5]) + MDC_SetMod.sStr(E_rec.E018[5]) + MDC_SetMod.sStr(E_rec.E019[5]) + MDC_SetMod.sStr(E_rec.E020[5]) + MDC_SetMod.sStr(E_rec.E021[5]) + MDC_SetMod.sStr(E_rec.E022[5]) + MDC_SetMod.sStr(E_rec.E023[5]) + MDC_SetMod.sStr(E_rec.E024[5]) + MDC_SetMod.sStr(E_rec.E025[5]) + MDC_SetMod.sStr(E_rec.E026[5]) + MDC_SetMod.sStr(E_rec.E027[5]) + MDC_SetMod.sStr(E_rec.E028[5]) + MDC_SetMod.sStr(E_rec.E029[5]) + MDC_SetMod.sStr(E_rec.E030[5]) + MDC_SetMod.sStr(E_rec.E031[5]) + MDC_SetMod.sStr(E_rec.E032[5]) + MDC_SetMod.sStr(E_rec.E033[5]) + MDC_SetMod.sStr(E_rec.E034[5]) + MDC_SetMod.sStr(E_rec.E035[5]) + MDC_SetMod.sStr(E_rec.E036[5]) + MDC_SetMod.sStr(E_rec.E037[5]) + MDC_SetMod.sStr(E_rec.E038[5]) + MDC_SetMod.sStr(E_rec.E039[5]) + MDC_SetMod.sStr(E_rec.E040[5]) + MDC_SetMod.sStr(E_rec.E041[5]) + MDC_SetMod.sStr(E_rec.E042[5]) + MDC_SetMod.sStr(E_rec.E043[5]) + MDC_SetMod.sStr(E_rec.E044[5]) + MDC_SetMod.sStr(E_rec.E045[5]) + MDC_SetMod.sStr(E_rec.E046[5]) + MDC_SetMod.sStr(E_rec.E047[5]) + MDC_SetMod.sStr(E_rec.E048[5]) + MDC_SetMod.sStr(E_rec.E049[5]) + MDC_SetMod.sStr(E_rec.E222));
//                    // / 다음줄넘김
//                    BUYCNT = 0;
//                    FAMCNT = FAMCNT + 1;
//                }
//            }
//        }
//        else
//            ErrNum = 1;

//        if (System.Convert.ToBoolean(CheckE) == false)
//            File_Create_E_record = true;
//        else
//            File_Create_E_record = false;
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
//        File_Create_E_record = false;
//    }

//    private bool File_Create_F_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 256619
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
//        short RCNT;
//        int iRow;

//        CheckF = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        RCNT = 1;
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / F_RECORE QUERY
//        sQry = "EXEC PH_PY980_F '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";
//        oRecordSet.DoQuery(sQry);

//        if (oRecordSet.RecordCount > 0)
//        {

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            F_rec.F001 = oRecordSet.Fields.Item("F001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            F_rec.F002 = oRecordSet.Fields.Item("F002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            F_rec.F003 = oRecordSet.Fields.Item("F003").VALUE;
//            F_rec.F004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(F_rec.F004))); // / C레코드의 일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            F_rec.F005 = oRecordSet.Fields.Item("F005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            F_rec.F006 = oRecordSet.Fields.Item("F006").VALUE;

//            SAVCNT = 0;
//            while (!oRecordSet.EOF)
//            {
//                SAVCNT = SAVCNT + 1; // / 레코드번호
//                                     // /초기화
//                if (SAVCNT == 1)
//                {
//                    for (iRow = 1; iRow <= 15; iRow++)
//                    {
//                        F_rec.F007[iRow] = Strings.Space(Strings.Len(F_rec.F007[iRow]));
//                        F_rec.F008[iRow] = Strings.Space(Strings.Len(F_rec.F008[iRow]));
//                        F_rec.F009[iRow] = Strings.Space(Strings.Len(F_rec.F009[iRow]));
//                        F_rec.F010[iRow] = Strings.Space(Strings.Len(F_rec.F010[iRow]));

//                        F_rec.F011[iRow] = VB6.Format(0, new string("0", Strings.Len(F_rec.F011[iRow])));
//                        F_rec.F012[iRow] = VB6.Format(0, new string("0", Strings.Len(F_rec.F012[iRow])));
//                        F_rec.F013[iRow] = VB6.Format(0, new string("0", Strings.Len(F_rec.F013[iRow])));
//                        F_rec.F014[iRow] = Strings.Space(Strings.Len(F_rec.F014[iRow]));
//                    }
//                }

//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                F_rec.F007[SAVCNT] = oRecordSet.Fields.Item("F007").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                F_rec.F008[SAVCNT] = oRecordSet.Fields.Item("F008").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                F_rec.F009[SAVCNT] = oRecordSet.Fields.Item("F009").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                F_rec.F010[SAVCNT] = oRecordSet.Fields.Item("F010").VALUE;
//                F_rec.F011[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("F011").VALUE, new string("0", Strings.Len(F_rec.F011[SAVCNT])));
//                F_rec.F012[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("F012").VALUE, new string("0", Strings.Len(F_rec.F012[SAVCNT])));
//                F_rec.F013[SAVCNT] = VB6.Format(oRecordSet.Fields.Item("F013").VALUE, new string("0", Strings.Len(F_rec.F013[SAVCNT])));
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                F_rec.F014[SAVCNT] = oRecordSet.Fields.Item("F014").VALUE;
//                oRecordSet.MoveNext();


//                if (SAVCNT == 15 | oRecordSet.EOF)
//                {
//                    F_rec.F127 = VB6.Format(RCNT, new string("0", Strings.Len(F_rec.F127))); // --일련번호
//                    F_rec.F128 = Strings.Space(Strings.Len(F_rec.F128));
//                    // / F레코드삽입
//                    PrintLine(1, MDC_SetMod.sStr(F_rec.F001) + MDC_SetMod.sStr(F_rec.F002) + MDC_SetMod.sStr(F_rec.F003) + MDC_SetMod.sStr(F_rec.F004) + MDC_SetMod.sStr(F_rec.F005) + MDC_SetMod.sStr(F_rec.F006) + MDC_SetMod.sStr(F_rec.F007[1]) + MDC_SetMod.sStr(F_rec.F008[1]) + MDC_SetMod.sStr(F_rec.F009[1]) + MDC_SetMod.sStr(F_rec.F010[1]) + MDC_SetMod.sStr(F_rec.F011[1]) + MDC_SetMod.sStr(F_rec.F012[1]) + MDC_SetMod.sStr(F_rec.F013[1]) + MDC_SetMod.sStr(F_rec.F014[1]) + MDC_SetMod.sStr(F_rec.F007[2]) + MDC_SetMod.sStr(F_rec.F008[2]) + MDC_SetMod.sStr(F_rec.F009[2]) + MDC_SetMod.sStr(F_rec.F010[2]) + MDC_SetMod.sStr(F_rec.F011[2]) + MDC_SetMod.sStr(F_rec.F012[2]) + MDC_SetMod.sStr(F_rec.F013[2]) + MDC_SetMod.sStr(F_rec.F014[2]) + MDC_SetMod.sStr(F_rec.F007[3]) + MDC_SetMod.sStr(F_rec.F008[3]) + MDC_SetMod.sStr(F_rec.F009[3]) + MDC_SetMod.sStr(F_rec.F010[3]) + MDC_SetMod.sStr(F_rec.F011[3]) + MDC_SetMod.sStr(F_rec.F012[3]) + MDC_SetMod.sStr(F_rec.F013[3]) + MDC_SetMod.sStr(F_rec.F014[3]) + MDC_SetMod.sStr(F_rec.F007[4]) + MDC_SetMod.sStr(F_rec.F008[4]) + MDC_SetMod.sStr(F_rec.F009[4]) + MDC_SetMod.sStr(F_rec.F010[4]) + MDC_SetMod.sStr(F_rec.F011[4]) + MDC_SetMod.sStr(F_rec.F012[4]) + MDC_SetMod.sStr(F_rec.F013[4]) + MDC_SetMod.sStr(F_rec.F014[4]) + MDC_SetMod.sStr(F_rec.F007[5]) + MDC_SetMod.sStr(F_rec.F008[5]) + MDC_SetMod.sStr(F_rec.F009[5]) + MDC_SetMod.sStr(F_rec.F010[5]) + MDC_SetMod.sStr(F_rec.F011[5]) + MDC_SetMod.sStr(F_rec.F012[5]) + MDC_SetMod.sStr(F_rec.F013[5]) + MDC_SetMod.sStr(F_rec.F014[5]) + MDC_SetMod.sStr(F_rec.F007[6]) + MDC_SetMod.sStr(F_rec.F008[6]) + MDC_SetMod.sStr(F_rec.F009[6]) + MDC_SetMod.sStr(F_rec.F010[6]) + MDC_SetMod.sStr(F_rec.F011[6]) + MDC_SetMod.sStr(F_rec.F012[6]) + MDC_SetMod.sStr(F_rec.F013[6]) + MDC_SetMod.sStr(F_rec.F014[6]) + MDC_SetMod.sStr(F_rec.F007[7]) + MDC_SetMod.sStr(F_rec.F008[7]) + MDC_SetMod.sStr(F_rec.F009[7]) + MDC_SetMod.sStr(F_rec.F010[7]) + MDC_SetMod.sStr(F_rec.F011[7]) + MDC_SetMod.sStr(F_rec.F012[7]) + MDC_SetMod.sStr(F_rec.F013[7]) + MDC_SetMod.sStr(F_rec.F014[7]) + MDC_SetMod.sStr(F_rec.F007[8]) + MDC_SetMod.sStr(F_rec.F008[8]) + MDC_SetMod.sStr(F_rec.F009[8]) + MDC_SetMod.sStr(F_rec.F010[8]) + MDC_SetMod.sStr(F_rec.F011[8]) + MDC_SetMod.sStr(F_rec.F012[8]) + MDC_SetMod.sStr(F_rec.F013[8]) + MDC_SetMod.sStr(F_rec.F014[8]) + MDC_SetMod.sStr(F_rec.F007[9]) + MDC_SetMod.sStr(F_rec.F008[9]) + MDC_SetMod.sStr(F_rec.F009[9]) + MDC_SetMod.sStr(F_rec.F010[9]) + MDC_SetMod.sStr(F_rec.F011[9]) + MDC_SetMod.sStr(F_rec.F012[9]) + MDC_SetMod.sStr(F_rec.F013[9]) + MDC_SetMod.sStr(F_rec.F014[9]) + MDC_SetMod.sStr(F_rec.F007[10]) + MDC_SetMod.sStr(F_rec.F008[10]) + MDC_SetMod.sStr(F_rec.F009[10]) + MDC_SetMod.sStr(F_rec.F010[10]) + MDC_SetMod.sStr(F_rec.F011[10]) + MDC_SetMod.sStr(F_rec.F012[10]) + MDC_SetMod.sStr(F_rec.F013[10]) + MDC_SetMod.sStr(F_rec.F014[10]) + MDC_SetMod.sStr(F_rec.F007[11]) + MDC_SetMod.sStr(F_rec.F008[11]) + MDC_SetMod.sStr(F_rec.F009[11]) + MDC_SetMod.sStr(F_rec.F010[11]) + MDC_SetMod.sStr(F_rec.F011[11]) + MDC_SetMod.sStr(F_rec.F012[11]) + MDC_SetMod.sStr(F_rec.F013[11]) + MDC_SetMod.sStr(F_rec.F014[11]) + MDC_SetMod.sStr(F_rec.F007[12]) + MDC_SetMod.sStr(F_rec.F008[12]) + MDC_SetMod.sStr(F_rec.F009[12]) + MDC_SetMod.sStr(F_rec.F010[12]) + MDC_SetMod.sStr(F_rec.F011[12]) + MDC_SetMod.sStr(F_rec.F012[12]) + MDC_SetMod.sStr(F_rec.F013[12]) + MDC_SetMod.sStr(F_rec.F014[12]) + MDC_SetMod.sStr(F_rec.F007[13]) + MDC_SetMod.sStr(F_rec.F008[13]) + MDC_SetMod.sStr(F_rec.F009[13]) + MDC_SetMod.sStr(F_rec.F010[13]) + MDC_SetMod.sStr(F_rec.F011[13]) + MDC_SetMod.sStr(F_rec.F012[13]) + MDC_SetMod.sStr(F_rec.F013[13]) + MDC_SetMod.sStr(F_rec.F014[13]) + MDC_SetMod.sStr(F_rec.F007[14]) + MDC_SetMod.sStr(F_rec.F008[14]) + MDC_SetMod.sStr(F_rec.F009[14]) + MDC_SetMod.sStr(F_rec.F010[14]) + MDC_SetMod.sStr(F_rec.F011[14]) + MDC_SetMod.sStr(F_rec.F012[14]) + MDC_SetMod.sStr(F_rec.F013[14]) + MDC_SetMod.sStr(F_rec.F014[14]) + MDC_SetMod.sStr(F_rec.F007[15]) + MDC_SetMod.sStr(F_rec.F008[15]) + MDC_SetMod.sStr(F_rec.F009[15]) + MDC_SetMod.sStr(F_rec.F010[15]) + MDC_SetMod.sStr(F_rec.F011[15]) + MDC_SetMod.sStr(F_rec.F012[15]) + MDC_SetMod.sStr(F_rec.F013[15]) + MDC_SetMod.sStr(F_rec.F014[15]) + MDC_SetMod.sStr(F_rec.F127) + MDC_SetMod.sStr(F_rec.F128));
//                    SAVCNT = 0;
//                    RCNT = RCNT + 1;
//                }
//            }
//        }


//        if (System.Convert.ToBoolean(CheckF) == false)
//            File_Create_F_record = true;
//        else
//            File_Create_F_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        Matrix_AddRow("F레코드오류: " + Information.Err.Source + " " + Information.Err.Description, ref false);
//        File_Create_F_record = false;
//    }

//    private bool File_Create_G_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 266229
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
//        string CheckG;

//        CheckG = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;
//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / G_RECORE QUERY
//        sQry = "EXEC PH_PY980_G '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";
//        oRecordSet.DoQuery(sQry);

//        if (oRecordSet.RecordCount > 0)
//        {

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G001 = oRecordSet.Fields.Item("G001").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G002 = oRecordSet.Fields.Item("G002").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G003 = oRecordSet.Fields.Item("G003").VALUE;
//            G_rec.G004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(G_rec.G004))); // / C레코드의 일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G005 = oRecordSet.Fields.Item("G005").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G006 = oRecordSet.Fields.Item("G006").VALUE;

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G007 = oRecordSet.Fields.Item("G007").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G008 = oRecordSet.Fields.Item("G008").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G009 = oRecordSet.Fields.Item("G009").VALUE;
//            G_rec.G010 = VB6.Format(oRecordSet.Fields.Item("G010").VALUE, new string("0", Strings.Len(G_rec.G010)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G011 = oRecordSet.Fields.Item("G011").VALUE;
//            G_rec.G012 = VB6.Format(oRecordSet.Fields.Item("G012").VALUE, new string("0", Strings.Len(G_rec.G012)));
//            G_rec.G013 = VB6.Format(oRecordSet.Fields.Item("G013").VALUE, new string("0", Strings.Len(G_rec.G013)));
//            G_rec.G014 = VB6.Format(oRecordSet.Fields.Item("G014").VALUE, new string("0", Strings.Len(G_rec.G014)));
//            G_rec.G015 = VB6.Format(oRecordSet.Fields.Item("G015").VALUE, new string("0", Strings.Len(G_rec.G015)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G016 = oRecordSet.Fields.Item("G016").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G017 = oRecordSet.Fields.Item("G017").VALUE;
//            G_rec.G018 = VB6.Format(oRecordSet.Fields.Item("G018").VALUE, new string("0", Strings.Len(G_rec.G018)));
//            G_rec.G019 = VB6.Format(oRecordSet.Fields.Item("G019").VALUE, new string("0", Strings.Len(G_rec.G019)));
//            G_rec.G020 = VB6.Format(oRecordSet.Fields.Item("G020").VALUE, new string("0", Strings.Len(G_rec.G020)));
//            G_rec.G021 = VB6.Format(oRecordSet.Fields.Item("G021").VALUE, new string("0", Strings.Len(G_rec.G021)));
//            G_rec.G022 = VB6.Format(oRecordSet.Fields.Item("G022").VALUE, new string("0", Strings.Len(G_rec.G022)));
//            G_rec.G023 = VB6.Format(oRecordSet.Fields.Item("G023").VALUE, new string("0", Strings.Len(G_rec.G023)));
//            G_rec.G024 = VB6.Format(oRecordSet.Fields.Item("G024").VALUE, new string("0", Strings.Len(G_rec.G024)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G025 = oRecordSet.Fields.Item("G025").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G026 = oRecordSet.Fields.Item("G026").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G027 = oRecordSet.Fields.Item("G027").VALUE;
//            G_rec.G028 = VB6.Format(oRecordSet.Fields.Item("G028").VALUE, new string("0", Strings.Len(G_rec.G028)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G029 = oRecordSet.Fields.Item("G029").VALUE;
//            G_rec.G030 = VB6.Format(oRecordSet.Fields.Item("G030").VALUE, new string("0", Strings.Len(G_rec.G030)));
//            G_rec.G031 = VB6.Format(oRecordSet.Fields.Item("G031").VALUE, new string("0", Strings.Len(G_rec.G031)));
//            G_rec.G032 = VB6.Format(oRecordSet.Fields.Item("G032").VALUE, new string("0", Strings.Len(G_rec.G032)));

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G033 = oRecordSet.Fields.Item("G033").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G034 = oRecordSet.Fields.Item("G034").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G035 = oRecordSet.Fields.Item("G035").VALUE;
//            G_rec.G036 = VB6.Format(oRecordSet.Fields.Item("G036").VALUE, new string("0", Strings.Len(G_rec.G036)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G037 = oRecordSet.Fields.Item("G037").VALUE;
//            G_rec.G038 = VB6.Format(oRecordSet.Fields.Item("G038").VALUE, new string("0", Strings.Len(G_rec.G038)));
//            G_rec.G039 = VB6.Format(oRecordSet.Fields.Item("G039").VALUE, new string("0", Strings.Len(G_rec.G039)));
//            G_rec.G040 = VB6.Format(oRecordSet.Fields.Item("G040").VALUE, new string("0", Strings.Len(G_rec.G040)));
//            G_rec.G041 = VB6.Format(oRecordSet.Fields.Item("G041").VALUE, new string("0", Strings.Len(G_rec.G041)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G042 = oRecordSet.Fields.Item("G042").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G043 = oRecordSet.Fields.Item("G043").VALUE;
//            G_rec.G044 = VB6.Format(oRecordSet.Fields.Item("G044").VALUE, new string("0", Strings.Len(G_rec.G044)));
//            G_rec.G045 = VB6.Format(oRecordSet.Fields.Item("G045").VALUE, new string("0", Strings.Len(G_rec.G045)));
//            G_rec.G046 = VB6.Format(oRecordSet.Fields.Item("G046").VALUE, new string("0", Strings.Len(G_rec.G046)));
//            G_rec.G047 = VB6.Format(oRecordSet.Fields.Item("G047").VALUE, new string("0", Strings.Len(G_rec.G047)));
//            G_rec.G048 = VB6.Format(oRecordSet.Fields.Item("G048").VALUE, new string("0", Strings.Len(G_rec.G048)));
//            G_rec.G049 = VB6.Format(oRecordSet.Fields.Item("G049").VALUE, new string("0", Strings.Len(G_rec.G049)));
//            G_rec.G050 = VB6.Format(oRecordSet.Fields.Item("G050").VALUE, new string("0", Strings.Len(G_rec.G050)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G051 = oRecordSet.Fields.Item("G051").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G052 = oRecordSet.Fields.Item("G052").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G053 = oRecordSet.Fields.Item("G053").VALUE;
//            G_rec.G054 = VB6.Format(oRecordSet.Fields.Item("G054").VALUE, new string("0", Strings.Len(G_rec.G054)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G055 = oRecordSet.Fields.Item("G055").VALUE;
//            G_rec.G056 = VB6.Format(oRecordSet.Fields.Item("G056").VALUE, new string("0", Strings.Len(G_rec.G056)));
//            G_rec.G057 = VB6.Format(oRecordSet.Fields.Item("G057").VALUE, new string("0", Strings.Len(G_rec.G057)));
//            G_rec.G058 = VB6.Format(oRecordSet.Fields.Item("G058").VALUE, new string("0", Strings.Len(G_rec.G058)));

//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G059 = oRecordSet.Fields.Item("G059").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G060 = oRecordSet.Fields.Item("G060").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G061 = oRecordSet.Fields.Item("G061").VALUE;
//            G_rec.G062 = VB6.Format(oRecordSet.Fields.Item("G062").VALUE, new string("0", Strings.Len(G_rec.G062)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G063 = oRecordSet.Fields.Item("G063").VALUE;
//            G_rec.G064 = VB6.Format(oRecordSet.Fields.Item("G064").VALUE, new string("0", Strings.Len(G_rec.G064)));
//            G_rec.G065 = VB6.Format(oRecordSet.Fields.Item("G065").VALUE, new string("0", Strings.Len(G_rec.G065)));
//            G_rec.G066 = VB6.Format(oRecordSet.Fields.Item("G066").VALUE, new string("0", Strings.Len(G_rec.G066)));
//            G_rec.G067 = VB6.Format(oRecordSet.Fields.Item("G067").VALUE, new string("0", Strings.Len(G_rec.G067)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G068 = oRecordSet.Fields.Item("G068").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G069 = oRecordSet.Fields.Item("G069").VALUE;
//            G_rec.G070 = VB6.Format(oRecordSet.Fields.Item("G070").VALUE, new string("0", Strings.Len(G_rec.G070)));
//            G_rec.G071 = VB6.Format(oRecordSet.Fields.Item("G071").VALUE, new string("0", Strings.Len(G_rec.G071)));
//            G_rec.G072 = VB6.Format(oRecordSet.Fields.Item("G072").VALUE, new string("0", Strings.Len(G_rec.G072)));
//            G_rec.G073 = VB6.Format(oRecordSet.Fields.Item("G073").VALUE, new string("0", Strings.Len(G_rec.G073)));
//            G_rec.G074 = VB6.Format(oRecordSet.Fields.Item("G074").VALUE, new string("0", Strings.Len(G_rec.G074)));
//            G_rec.G075 = VB6.Format(oRecordSet.Fields.Item("G075").VALUE, new string("0", Strings.Len(G_rec.G075)));
//            G_rec.G076 = VB6.Format(oRecordSet.Fields.Item("G076").VALUE, new string("0", Strings.Len(G_rec.G076)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G077 = oRecordSet.Fields.Item("G077").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G078 = oRecordSet.Fields.Item("G078").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G079 = oRecordSet.Fields.Item("G079").VALUE;
//            G_rec.G080 = VB6.Format(oRecordSet.Fields.Item("G080").VALUE, new string("0", Strings.Len(G_rec.G080)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G081 = oRecordSet.Fields.Item("G081").VALUE;
//            G_rec.G082 = VB6.Format(oRecordSet.Fields.Item("G082").VALUE, new string("0", Strings.Len(G_rec.G082)));
//            G_rec.G083 = VB6.Format(oRecordSet.Fields.Item("G083").VALUE, new string("0", Strings.Len(G_rec.G083)));
//            G_rec.G084 = VB6.Format(oRecordSet.Fields.Item("G084").VALUE, new string("0", Strings.Len(G_rec.G084)));
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G085 = oRecordSet.Fields.Item("G085").VALUE;
//            // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            G_rec.G086 = oRecordSet.Fields.Item("G086").VALUE;


//            PrintLine(1, MDC_SetMod.sStr(G_rec.G001) + MDC_SetMod.sStr(G_rec.G002) + MDC_SetMod.sStr(G_rec.G003) + MDC_SetMod.sStr(G_rec.G004) + MDC_SetMod.sStr(G_rec.G005) + MDC_SetMod.sStr(G_rec.G006) + MDC_SetMod.sStr(G_rec.G007) + MDC_SetMod.sStr(G_rec.G008) + MDC_SetMod.sStr(G_rec.G009) + MDC_SetMod.sStr(G_rec.G010) + MDC_SetMod.sStr(G_rec.G011) + MDC_SetMod.sStr(G_rec.G012) + MDC_SetMod.sStr(G_rec.G013) + MDC_SetMod.sStr(G_rec.G014) + MDC_SetMod.sStr(G_rec.G015) + MDC_SetMod.sStr(G_rec.G016) + MDC_SetMod.sStr(G_rec.G017) + MDC_SetMod.sStr(G_rec.G018) + MDC_SetMod.sStr(G_rec.G019) + MDC_SetMod.sStr(G_rec.G020) + MDC_SetMod.sStr(G_rec.G021) + MDC_SetMod.sStr(G_rec.G022) + MDC_SetMod.sStr(G_rec.G023) + MDC_SetMod.sStr(G_rec.G024) + MDC_SetMod.sStr(G_rec.G025) + MDC_SetMod.sStr(G_rec.G026) + MDC_SetMod.sStr(G_rec.G027) + MDC_SetMod.sStr(G_rec.G028) + MDC_SetMod.sStr(G_rec.G029) + MDC_SetMod.sStr(G_rec.G030) + MDC_SetMod.sStr(G_rec.G031) + MDC_SetMod.sStr(G_rec.G032) + MDC_SetMod.sStr(G_rec.G033) + MDC_SetMod.sStr(G_rec.G034) + MDC_SetMod.sStr(G_rec.G035) + MDC_SetMod.sStr(G_rec.G036) + MDC_SetMod.sStr(G_rec.G037) + MDC_SetMod.sStr(G_rec.G038) + MDC_SetMod.sStr(G_rec.G039) + MDC_SetMod.sStr(G_rec.G040) + MDC_SetMod.sStr(G_rec.G041) + MDC_SetMod.sStr(G_rec.G042) + MDC_SetMod.sStr(G_rec.G043) + MDC_SetMod.sStr(G_rec.G044) + MDC_SetMod.sStr(G_rec.G045) + MDC_SetMod.sStr(G_rec.G046) + MDC_SetMod.sStr(G_rec.G047) + MDC_SetMod.sStr(G_rec.G048) + MDC_SetMod.sStr(G_rec.G049) + MDC_SetMod.sStr(G_rec.G050) + MDC_SetMod.sStr(G_rec.G051) + MDC_SetMod.sStr(G_rec.G052) + MDC_SetMod.sStr(G_rec.G053) + MDC_SetMod.sStr(G_rec.G054) + MDC_SetMod.sStr(G_rec.G055) + MDC_SetMod.sStr(G_rec.G056) + MDC_SetMod.sStr(G_rec.G057) + MDC_SetMod.sStr(G_rec.G058) + MDC_SetMod.sStr(G_rec.G059) + MDC_SetMod.sStr(G_rec.G060) + MDC_SetMod.sStr(G_rec.G061) + MDC_SetMod.sStr(G_rec.G062) + MDC_SetMod.sStr(G_rec.G063) + MDC_SetMod.sStr(G_rec.G064) + MDC_SetMod.sStr(G_rec.G065) + MDC_SetMod.sStr(G_rec.G066) + MDC_SetMod.sStr(G_rec.G067) + MDC_SetMod.sStr(G_rec.G068) + MDC_SetMod.sStr(G_rec.G069) + MDC_SetMod.sStr(G_rec.G070) + MDC_SetMod.sStr(G_rec.G071) + MDC_SetMod.sStr(G_rec.G072) + MDC_SetMod.sStr(G_rec.G073) + MDC_SetMod.sStr(G_rec.G074) + MDC_SetMod.sStr(G_rec.G075) + MDC_SetMod.sStr(G_rec.G076) + MDC_SetMod.sStr(G_rec.G077) + MDC_SetMod.sStr(G_rec.G078) + MDC_SetMod.sStr(G_rec.G079) + MDC_SetMod.sStr(G_rec.G080) + MDC_SetMod.sStr(G_rec.G081) + MDC_SetMod.sStr(G_rec.G082) + MDC_SetMod.sStr(G_rec.G083) + MDC_SetMod.sStr(G_rec.G084) + MDC_SetMod.sStr(G_rec.G085) + MDC_SetMod.sStr(G_rec.G086));
//        }


//        if (System.Convert.ToBoolean(CheckG) == false)
//            File_Create_G_record = true;
//        else
//            File_Create_G_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;

//        Matrix_AddRow("G레코드오류: " + Information.Err.Source + " " + Information.Err.Description, ref false);
//        File_Create_G_record = false;
//    }

//    private bool File_Create_H_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 284295
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
//        string CheckH;
//        short JONCNT;
//        short HCount;

//        CheckH = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / H_RECORE QUERY
//        sQry = "EXEC PH_PY980_H '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";
//        oRecordSet.DoQuery(sQry);

//        HCount = 0;

//        if (oRecordSet.RecordCount > 0)
//        {
//            while (!oRecordSet.EOF)
//            {

//                // H RECORD MOVE

//                HCount = HCount + 1;

//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H001 = oRecordSet.Fields.Item("H001").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H002 = oRecordSet.Fields.Item("H002").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H003 = oRecordSet.Fields.Item("H003").VALUE;
//                H_rec.H004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(H_rec.H004))); // / 일련번호
//                                                                                               // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H005 = oRecordSet.Fields.Item("H005").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H006 = oRecordSet.Fields.Item("H006").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H007 = oRecordSet.Fields.Item("H007").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H008 = oRecordSet.Fields.Item("H008").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H009 = oRecordSet.Fields.Item("H009").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H010 = oRecordSet.Fields.Item("H010").VALUE;

//                H_rec.H011 = VB6.Format(oRecordSet.Fields.Item("H011").VALUE, new string("0", Strings.Len(H_rec.H011)));
//                H_rec.H012 = VB6.Format(oRecordSet.Fields.Item("H012").VALUE, new string("0", Strings.Len(H_rec.H012)));
//                H_rec.H013 = VB6.Format(oRecordSet.Fields.Item("H013").VALUE, new string("0", Strings.Len(H_rec.H013)));
//                H_rec.H014 = VB6.Format(oRecordSet.Fields.Item("H014").VALUE, new string("0", Strings.Len(H_rec.H014)));
//                H_rec.H015 = VB6.Format(oRecordSet.Fields.Item("H015").VALUE, new string("0", Strings.Len(H_rec.H015)));
//                H_rec.H016 = VB6.Format(oRecordSet.Fields.Item("H016").VALUE, new string("0", Strings.Len(H_rec.H016)));
//                H_rec.H017 = VB6.Format(oRecordSet.Fields.Item("H017").VALUE, new string("0", Strings.Len(H_rec.H017)));
//                H_rec.H018 = VB6.Format(HCount, new string("0", Strings.Len(H_rec.H018))); // / 일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                H_rec.H019 = oRecordSet.Fields.Item("H019").VALUE;

//                PrintLine(1, MDC_SetMod.sStr(H_rec.H001) + MDC_SetMod.sStr(H_rec.H002) + MDC_SetMod.sStr(H_rec.H003) + MDC_SetMod.sStr(H_rec.H004) + MDC_SetMod.sStr(H_rec.H005) + MDC_SetMod.sStr(H_rec.H006) + MDC_SetMod.sStr(H_rec.H007) + MDC_SetMod.sStr(H_rec.H008) + MDC_SetMod.sStr(H_rec.H009) + MDC_SetMod.sStr(H_rec.H010) + MDC_SetMod.sStr(H_rec.H011) + MDC_SetMod.sStr(H_rec.H012) + MDC_SetMod.sStr(H_rec.H013) + MDC_SetMod.sStr(H_rec.H014) + MDC_SetMod.sStr(H_rec.H015) + MDC_SetMod.sStr(H_rec.H016) + MDC_SetMod.sStr(H_rec.H017) + MDC_SetMod.sStr(H_rec.H018) + MDC_SetMod.sStr(H_rec.H019));
//                oRecordSet.MoveNext();
//            }
//        }

//        if (System.Convert.ToBoolean(CheckH) == false)
//            File_Create_H_record = true;
//        else
//            File_Create_H_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        Matrix_AddRow("H레코드오류: " + Information.Err.Description, ref false);
//        File_Create_H_record = false;
//    }

//    private bool File_Create_I_record()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo Error_Message' at character 289737
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
//        string CheckI;
//        short JONCNT;
//        short ICount;

//        CheckI = System.Convert.ToString(false); // /체크필요유무
//        ErrNum = 0;

//        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        // / I_RECORE QUERY
//        sQry = "EXEC PH_PY980_I '" + C_SAUP + "', '" + C_YYYY + "', '" + C_SABUN + "'";
//        oRecordSet.DoQuery(sQry);

//        ICount = 0;

//        if (oRecordSet.RecordCount > 0)
//        {
//            while (!oRecordSet.EOF)
//            {

//                // I RECORD MOVE

//                ICount = ICount + 1;

//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I001 = oRecordSet.Fields.Item("I001").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I002 = oRecordSet.Fields.Item("I002").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I003 = oRecordSet.Fields.Item("I003").VALUE;
//                I_rec.I004 = VB6.Format(C_rec.C004, new string("0", Strings.Len(I_rec.I004))); // / C레코드의 일련번호
//                                                                                               // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I005 = oRecordSet.Fields.Item("I005").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I006 = oRecordSet.Fields.Item("I006").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I007 = oRecordSet.Fields.Item("I007").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I008 = oRecordSet.Fields.Item("I008").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I009 = oRecordSet.Fields.Item("I009").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I010 = oRecordSet.Fields.Item("I010").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I011 = oRecordSet.Fields.Item("I011").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I012 = oRecordSet.Fields.Item("I012").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I013 = oRecordSet.Fields.Item("I013").VALUE;
//                // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I014 = oRecordSet.Fields.Item("I014").VALUE;
//                I_rec.I015 = VB6.Format(oRecordSet.Fields.Item("I015").VALUE, new string("0", Strings.Len(I_rec.I015)));
//                I_rec.I016 = VB6.Format(oRecordSet.Fields.Item("I016").VALUE, new string("0", Strings.Len(I_rec.I016)));
//                I_rec.I017 = VB6.Format(oRecordSet.Fields.Item("I017").VALUE, new string("0", Strings.Len(I_rec.I017)));
//                I_rec.I018 = VB6.Format(oRecordSet.Fields.Item("I018").VALUE, new string("0", Strings.Len(I_rec.I018)));
//                I_rec.I019 = VB6.Format(oRecordSet.Fields.Item("I019").VALUE, new string("0", Strings.Len(I_rec.I019)));
//                I_rec.I020 = VB6.Format(ICount, new string("0", Strings.Len(I_rec.I020))); // /일련번호
//                                                                                           // UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                I_rec.I021 = oRecordSet.Fields.Item("I021").VALUE;

//                PrintLine(1, MDC_SetMod.sStr(I_rec.I001) + MDC_SetMod.sStr(I_rec.I002) + MDC_SetMod.sStr(I_rec.I003) + MDC_SetMod.sStr(I_rec.I004) + MDC_SetMod.sStr(I_rec.I005) + MDC_SetMod.sStr(I_rec.I006) + MDC_SetMod.sStr(I_rec.I007) + MDC_SetMod.sStr(I_rec.I008) + MDC_SetMod.sStr(I_rec.I009) + MDC_SetMod.sStr(I_rec.I010) + MDC_SetMod.sStr(I_rec.I011) + MDC_SetMod.sStr(I_rec.I012) + MDC_SetMod.sStr(I_rec.I013) + MDC_SetMod.sStr(I_rec.I014) + MDC_SetMod.sStr(I_rec.I015) + MDC_SetMod.sStr(I_rec.I016) + MDC_SetMod.sStr(I_rec.I017) + MDC_SetMod.sStr(I_rec.I018) + MDC_SetMod.sStr(I_rec.I019) + MDC_SetMod.sStr(I_rec.I020) + MDC_SetMod.sStr(I_rec.I021));
//                oRecordSet.MoveNext();
//            }
//        }
//        if (System.Convert.ToBoolean(CheckI) == false)
//            File_Create_I_record = true;
//        else
//            File_Create_I_record = false;
//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        return;
//        // ///////////////////////////////////////////////////////////////////////////////////////////////////////
//        Error_Message:
//        ;

//        // UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        oRecordSet = null/* TODO Change to default(_) if this is not a reference type */;
//        Matrix_AddRow("I레코드오류: " + Information.Err.Description, ref false);
//        File_Create_I_record = false;
//    }

//    private void Matrix_AddRow(string MatrixMsg, ref bool Insert_YN = false, ref bool MatrixErr = false)
//    {
//        if (MatrixErr == true)
//            oForm.DataSources.UserDataSources.Item("Col0").VALUE = "??";
//        else
//            oForm.DataSources.UserDataSources.Item("Col0").VALUE = "";
//        oForm.DataSources.UserDataSources.Item("Col1").VALUE = MatrixMsg;
//        if (Insert_YN == true)
//        {
//            oMat1.AddRow();
//            MaxRow = MaxRow + 1;
//        }
//        oMat1.SetLineData(MaxRow);
//    }
//    // 화면변수 CHECK
//    private bool HeaderSpaceLineDel()
//    {
//        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo HeaderSpaceLi...' at character 296551
//   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
//   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
//   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
//   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

//Input: 
//		On Error GoTo HeaderSpaceLineDel

// */		short ErrNum;

//        ErrNum = 0;
//        // / 필수Check
//        // UPGRADE_WARNING: oForm.Items(HtaxID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        if (oForm.Items.Item("HtaxID").Specific.VALUE == "")
//        {
//            ErrNum = 1;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("TeamName").Specific.VALUE == "")
//        {
//            ErrNum = 2;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("Dname").Specific.VALUE == "")
//        {
//            ErrNum = 3;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("Dtel").Specific.VALUE == "")
//        {
//            ErrNum = 4;
//            goto HeaderSpaceLineDel;
//        }
//        else if (oForm.Items.Item("DocDate").Specific.VALUE == "")
//        {
//            ErrNum = 5;
//            goto HeaderSpaceLineDel;
//        }

//        HeaderSpaceLineDel = true;
//        return;
//        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        HeaderSpaceLineDel:
//        ;
//        if (ErrNum == 1)
//            Sbo_Application.StatusBar.SetText("홈텍스ID(5자리이상)를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 2)
//            Sbo_Application.StatusBar.SetText("담당자부서는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 3)
//            Sbo_Application.StatusBar.SetText("담당자성명은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 4)
//            Sbo_Application.StatusBar.SetText("담당자전화번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else if (ErrNum == 5)
//            Sbo_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        else
//            Sbo_Application.StatusBar.SetText("HeaderSpaceLineDel 실행 중 오류가 발생했습니다." + Strings.Space(10) + Information.Err.Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

//        HeaderSpaceLineDel = false;
//    }
//}
