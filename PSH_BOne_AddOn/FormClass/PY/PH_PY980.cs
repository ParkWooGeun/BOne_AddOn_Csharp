using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 근로소득지급명세서자료 전산매체수록
    /// </summary>
    internal class PH_PY980 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        
        /// <summary>
        /// Form 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
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
        /// 신고파일 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create()
        {
            bool returnValue = false;
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
                returnValue = true;
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
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// A레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_A_record(string pCLTCOD, string pyyyy, string pHtaxID, string pTeamName, string pDname, string pDtel, string pDocDate)
        {
            bool returnValue = false;
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
            // 2020년귀속 1893 BYTE

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
            string A017; // 1691  '공란

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
                    A017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A017").Value.ToString().Trim(), 1691);

                    FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015 + A016 + A017);

                }

                returnValue = true;
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

            return returnValue;
        }

        /// <summary>
        /// B레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_B_record(string pCLTCOD, string pyyyy)
        {
            bool returnValue = false;
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
            string B018; // 1683  '공란

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
                    B018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B018").Value.ToString().Trim(), 1683);

                    FileSystem.PrintLine(1, B001 + B002 + B003 + B004 + B005 + B006 + B007 + B008 + B009 + B010 + B011 + B012 + B013 + B014 + B015 + B016 + B017 + B018);

                }

                returnValue = true;
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

            return returnValue;
        }

        /// <summary>
        /// C레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_C_record(string pCLTCOD, string pyyyy)
        {
            bool returnValue = false;
            short errNum = 0;
            string sQry = string.Empty;
            string c_SAUP = string.Empty;
            string c_YYYY = string.Empty;
            string c_SABUN = string.Empty;
            int newCNT = 0; //일련번호

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
            string C033;    // 11    '공란
            string C034;    // 11    '공란
            string C035;    // 11    '계
            // 주(현)근무처 비과세 및 감면소득
            string C036;    // 10    '비과세(G01:학자금)
            string C037;    // 10    '비과세(H01:무보수위원수당)
            string C038;    // 10    '비과세(H05:경호,승선수당)
            string C039;    // 10    '비과세(H06:유아,초중등)
            string C040;    // 10    '비과세(H07:고등교육법)
            string C041;    // 10    '비과세(H08:특별법)
            string C042;    // 10    '비과세(H10:기업부설연구소)
            string C043;    // 10    '비과세(H09:연구기관등)
            string C044;    // 10    '비과세(H14:보육교사근무환경개선비)
            string C045;    // 10    '비과세(H15:사립유치원수석교사.교사의인건비)
            string C046;    // 10    '비과세(H11:취재수당)
            string C047;    // 10    '비과세(H12:벽지수당)
            string C048;    // 10    '비과세(H13:재해관련급여)
            string C049;    // 10    '비과세(H16:정부공공기관지방이전기관종사자이주수당)
            string C050;    // 10    '비과세(H17:종교활동비)
            string C051;    // 10    '비과세(I01:외국정부등근로자)
            string C052;    // 10    '비과세(K01:외국주둔군인등)
            string C053;    // 10    '비과세(M01:국외근로100만원)
            string C054;    // 10    '비과세(M02:국외근로300만원)
            string C055;    // 10    '비과세(M03:국외근로)
            string C056;    // 10    '비과세(O01:야간근로수당)
            string C057;    // 10    '비과세(Q01:출산보육수당)
            string C058;    // 10    '비과세(R10:근로장학금)
            string C059;    // 10    '비과세(R11:직무발명보상금)
            string C060;    // 10    '비과세(S01:주식매수선택권)
            string C061;    // 10    '비과세(U01:벤처기업주식매수선택권)
            string C062A;   // 10    '비과세(Y02:우리사주조합인출금50%)
            string C062B;   // 10    '비과세(Y03:우리사주조합인출금75%)
            string C062C;   // 10    '비과세(Y03:우리사주조합인출금100%)
            string C063;    // 10    '비과세(Y22:전공의수련보조수당)
            string C064;    // 10    '비과세(T01:외국인기술자)
            string C065;    // 10    '비과세(T30:성과공유중소기업경영성과급)
            string C066;    // 10    '비과세(T40:중소기업핵심인력성과보상기금소득세감면)
            string C067;    // 10    '비과세(T50:내국인우수인력국내복귀소득세감면)
            string C068A;   // 10    '비과세(T11:중소기업취업청년소득세감면50%)
            string C068B;   // 10    '비과세(T12:중소기업취업청년소득세감면70%)
            string C068C;   // 10    '비과세(T13:중소기업취업청년소득세감면90%)
            string C069;    // 10    '비과세(T20:조세조약상교직자감면)
            string C070;    // 10    '공란
            string C071;    // 10    '공란
            string C072;    // 10    '공란
            string C073;    // 10    '공란
            string C074;    // 10    '비과세 계
            string C075;    // 10    '감면소득 계
            // 정산명세    
            string C076;    // 11    '총급여
            string C077;    // 10    '근로소득공제
            string C078;    // 11    '근로소득금액
            // 기본공제    
            string C079;    // 8     '본인공제금액
            string C080;    // 8     '배우자공제금액
            string C081A;   // 2     '부양가족공제인원
            string C081B;   // 8     '부양가족공제금액
            // 추가공제  
            string C082A;   // 2     '경로우대공제인원
            string C082B;   // 8     '경로우대공제금액
            string C083A;   // 2     '장애자공제인원
            string C083B;   // 8     '장애자공제금액
            string C084;    // 8     '부녀자공제금액
            string C085;    // 10    '한부모공제금액
            // 연금보험료공
            string C086A;   // 10    '국민연금보험료공제_대상금액
            string C086B;   // 10    '국민연금보험료공제_공제금액
            string C087A;   // 10    '공적연금보험료공제_공무원연금_대상금액
            string C087B;   // 10    '공적연금보험료공제_공무원연금_공제금액
            string C088A;   // 10    '공적연금보험료공제_군인연금_대상금액
            string C088B;   // 10    '공적연금보험료공제_군인연금_공제금액
            string C089A;   // 10    '공적연금보험료공제_사립학교교직원연금_대상금액
            string C089B;   // 10    '공적연금보험료공제_립학교교직원연금_공제금액
            string C090A;   // 10    '공적연금보험료공제_별정우체국연금_대상금액
            string C090B;   // 10    '공적연금보험료공제_별정우체국연금_공제금액
            // 특별소득공제
            string C091A;   // 10    '보험료_건강보험료_대상금액
            string C091B;   // 10    '보험료_건강보험료_공제금액
            string C092A;   // 10    '보험료_고용보험료_대상금액
            string C092B;   // 10    '보험료_고용보험료_공제금액
            string C093A;   // 8     '주택자금_주택임차차입금 원리금상환공제금액-대출기관
            string C093B;   // 8     '주택자금_주택임차차입금 원리금상환공제금액-거주자
            string C094A;   // 8     '2011 장기주택저당차입금이자상환공제금액-15년미만
            string C094B;   // 8     '2011 장기주택저당차입금이자상환공제금액-15-29년
            string C094C;   // 8     '2011 장기주택저당차입금이자상환공제금액-30년이상
            string C095A;   // 8     '2012 이후차입분,15년이상-고정금리비거치상환대출
            string C095B;   // 8     '2012 이후차입분,15년이상-기타대출
            string C096A;   // 8     '2015 이후차입분,15년이상-고정금리이면서비거치상환대출
            string C096B;   // 8     '2015 이후차입분,15년이상-고정금리이거나비거치상환대출
            string C096C;   // 8     '2015 이후차입분,15년이상-기타대출
            string C096D;   // 8     '2015 이후차입분,10~15년-고정금리이거나비거치상환대출
            string C097;    // 11    '기부금(이월분)
            string C098;    // 11    '공란
            string C099;    // 11    '공란
            string C100;    // 11    '계  특별소득공제계
            string C101;    // 11    '차감소득금액
            // 그밖의소득공제
            string C102;    // 8     '개인연금저축소득공제
            string C103;    // 10    '소기업소상공인공제부금
            string C104;    // 10    '주택마련저축소득공제_청약저축
            string C105;    // 10    '주택마련저축소득공제_주택청약종합저축
            string C106;    // 10    '주택마련저축소득공제_근로자주택마련저축
            string C107;    // 10    '투자조합출자등소득공제
            string C108;    // 8     '신용카드등소득공제
            string C109;    // 10    '우리사주조합출연금
            string C110;    // 10    '고용유지중소기업근로자소득공제
            string C111;    // 10    '장기집합투자증권저축
            string C112;    // 10    '공란 '0'
            string C113;    // 10    '공란 '0'
            string C114;    // 11    '그밖의소득공제계
            string C115;    // 11    '소득공제종합한도초과액
            string C116;    // 11    '종합소득과세표준
            string C117;    // 11    '산출세액
            // 세액감면     
            string C118;    // 10    '소득세법
            string C119;    // 10    '조특법
            string C120;    // 10    '조특법제30조
            string C121;    // 10    '조세조약
            string C122;    // 10    '공란
            string C123;    // 10    '공란
            string C124;    // 10    '세액감면계
            // 세액공제
            string C125;    // 10    '근로소득세액공제
            string C126A;   // 2     '자녀세액공제인원
            string C126B;   // 10    '자녀세액공제
            string C127A;   // 2     '출산.입양세액공제인원
            string C127B;   // 10    '출산.입양세액공제
            string C128A;   // 10    '연금계좌_과학기술인공제_공제대상금액
            string C128B;   // 10    '연금계좌_과학기술인공제_세액공제액
            string C129A;   // 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_공제대상금액
            string C129B;   // 10    '연금계좌_근로자퇴직급여보장법에따른 퇴직급여_세액공제액
            string C130A;   // 10    '연금계좌_연금저축_공제대상금액
            string C130B;   // 10    '연금계좌_연금저축_세액공제액
            string C131A;   // 10    '특별세액공제_보장성보험료_공제대상금액
            string C131B;   // 10    '특별세액공제_보장성보험료_세액공제액
            string C132A;   // 10    '특별세액공제_장애인전용보험료_공제대상금액
            string C132B;   // 10    '특별세액공제_장애인전용보험료_세액공제액
            string C133A;   // 10    '특별세액공제_의료비_공제대상금액
            string C133B;   // 10    '특별세액공제_의료비_세액공제액
            string C134A;   // 10    '특별세액공제_교육비_공제대상금액
            string C134B;   // 10    '특별세액공제_교육비_세액공제액
            string C135A;   // 10    '특별세액공제_기부금_정치자금_10만원이하_공제대상금액
            string C135B;   // 10    '특별세액공제_기부금_정치자금_10만원이하_세액공제액
            string C136A;   // 11    '특별세액공제_기부금_정치자금_10만원초과_공제대상금액
            string C136B;   // 10    '특별세액공제_기부금_정치자금_10만원초과_세액공제액
            string C137A;   // 11    '특별세액공제_기부금_법정기부금_공제대상금액
            string C137B;   // 10    '특별세액공제_기부금_법정기부금_세액공제액
            string C138A;   // 11    '특별세액공제_기부금_우리사주조합기부금_공제대상금액
            string C138B;   // 10    '특별세액공제_기부금_우리사주조합기부금_세액공제액
            string C139A;   // 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체외)
            string C139B;   // 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체외)
            string C140A;   // 11    '특별세액공제_기부금_지정기부금_공제대상금액(종교단체)
            string C140B;   // 10    '특별세액공제_기부금_지정기부금_세액공제액(종교단체)
            string C141;    // 11    '공란 '0'
            string C142;    // 11    '공란 '0'
            string C143;    // 10    '특별세액공제계
            string C144;    // 10    '표준세액공제
            string C145;    // 10    '납세조합공제
            string C146;    // 10    '주택차입금
            string C147;    // 10    '외국납부
            string C148A;   // 10    '월세세액공제_공제대상금액
            string C148B;   // 8     '월세세액공제_세액공제액
            string C149;    // 10    '공란 '0'
            string C150;    // 10    '공란 '0'
            string C151;    // 10    '세액공제계
            // 결정세액
            string C152A;   // 11    '소득세
            string C152B;   // 10    '지방소득세
            string C152C;   // 10    '농특세
            string C153;    // 3     '실효세율
            // 기납부세액_주(현)근무지
            string C154A;   // 11    '소득세
            string C154B;   // 10    '지방소득세
            string C154C;   // 10    '농특세
            // 납부특례세액
            string C155A;   // 11    '소득세
            string C155B;   // 10    '지방소득세
            string C155C;   // 10    '농특세
            // 차감징수세액
            string C156A_1; // 1    '소득세(기호 음수1, 양수0)
            string C156A_2; // 11   '소득세
            string C156B_1; // 1    '지방소득세(기호 음수1, 양수0)
            string C156B_2; // 10   '지방소득세
            string C156C_1; // 1    '농특세(기호 음수1, 양수0)
            string C156C_2; // 10   '농특세

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
                    newCNT = 0;
                    while (!oRecordSet.EoF)
                    {
                        newCNT += 1; //일련번호
                        c_SAUP = oRecordSet.Fields.Item("saup").Value.ToString().Trim();
                        c_YYYY = oRecordSet.Fields.Item("yyyy").Value.ToString().Trim();
                        c_SABUN = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();

                        //C RECORD MOVE
                        C001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C001").Value.ToString().Trim(), 1);
                        C002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C002").Value.ToString().Trim(), 2);
                        C003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C003").Value.ToString().Trim(), 3);
                        C004 = codeHelpClass.GetFixedLengthStringByte(newCNT.ToString(), 6, '0');
                        C005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C005").Value.ToString().Trim(), 10);
                        C006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C006").Value.ToString().Trim(), 2, '0');
                        C007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C007").Value.ToString().Trim(), 1);
                        C008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C008").Value.ToString().Trim(), 2);
                        C009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C009").Value.ToString().Trim(), 1);
                        C010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C010").Value.ToString().Trim(), 1);

                        C011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C011").Value.ToString().Trim(), 30);
                        C012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C012").Value.ToString().Trim(), 1);
                        C013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C013").Value.ToString().Trim(), 13);
                        C014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C014").Value.ToString().Trim(), 2);
                        C015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C015").Value.ToString().Trim(), 1);
                        C016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C016").Value.ToString().Trim(), 1);
                        C017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C017").Value.ToString().Trim(), 1);
                        C018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C018").Value.ToString().Trim(), 4);
                        C019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C019").Value.ToString().Trim(), 1);
                        C020 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C020").Value.ToString().Trim(), 10);

                        C021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C021").Value.ToString().Trim(), 60);
                        C022 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C022").Value.ToString().Trim(), 8, '0');
                        C023 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C023").Value.ToString().Trim(), 8, '0');
                        C024 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C024").Value.ToString().Trim(), 8, '0');
                        C025 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C025").Value.ToString().Trim(), 8, '0');
                        C026 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C026").Value.ToString().Trim(), 11, '0');
                        C027 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C027").Value.ToString().Trim(), 11, '0');
                        C028 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C028").Value.ToString().Trim(), 11, '0');
                        C029 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C029").Value.ToString().Trim(), 11, '0');
                        C030 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C030").Value.ToString().Trim(), 11, '0');

                        C031 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C031").Value.ToString().Trim(), 11, '0');
                        C032 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C032").Value.ToString().Trim(), 11, '0');
                        C033 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C033").Value.ToString().Trim(), 11, '0');
                        C034 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C034").Value.ToString().Trim(), 11, '0');
                        C035 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C035").Value.ToString().Trim(), 11, '0');
                        C036 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C036").Value.ToString().Trim(), 10, '0');
                        C037 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C037").Value.ToString().Trim(), 10, '0');
                        C038 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C038").Value.ToString().Trim(), 10, '0');
                        C039 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C039").Value.ToString().Trim(), 10, '0');
                        C040 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C040").Value.ToString().Trim(), 10, '0');

                        C041 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C041").Value.ToString().Trim(), 10, '0');
                        C042 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C042").Value.ToString().Trim(), 10, '0');
                        C043 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C043").Value.ToString().Trim(), 10, '0');
                        C044 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C044").Value.ToString().Trim(), 10, '0');
                        C045 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C045").Value.ToString().Trim(), 10, '0');
                        C046 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C046").Value.ToString().Trim(), 10, '0');
                        C047 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C047").Value.ToString().Trim(), 10, '0');
                        C048 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C048").Value.ToString().Trim(), 10, '0');
                        C049 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C049").Value.ToString().Trim(), 10, '0');
                        C050 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C050").Value.ToString().Trim(), 10, '0');

                        C051 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C051").Value.ToString().Trim(), 10, '0');
                        C052 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C052").Value.ToString().Trim(), 10, '0');
                        C053 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C053").Value.ToString().Trim(), 10, '0');
                        C054 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C054").Value.ToString().Trim(), 10, '0');
                        C055 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C055").Value.ToString().Trim(), 10, '0');
                        C056 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C056").Value.ToString().Trim(), 10, '0');
                        C057 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C057").Value.ToString().Trim(), 10, '0');
                        C058 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C058").Value.ToString().Trim(), 10, '0');
                        C059 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C059").Value.ToString().Trim(), 10, '0');
                        C060 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C060").Value.ToString().Trim(), 10, '0');

                        C061 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C061").Value.ToString().Trim(), 10, '0');
                        C062A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C062A").Value.ToString().Trim(), 10, '0');
                        C062B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C062B").Value.ToString().Trim(), 10, '0');
                        C062C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C062C").Value.ToString().Trim(), 10, '0');
                        C063 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C063").Value.ToString().Trim(), 10, '0');
                        C064 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C064").Value.ToString().Trim(), 10, '0');
                        C065 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C065").Value.ToString().Trim(), 10, '0');
                        C066 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C066").Value.ToString().Trim(), 10, '0');
                        C067 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C067").Value.ToString().Trim(), 10, '0');
                        C068A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C068A").Value.ToString().Trim(), 10, '0');
                        C068B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C068B").Value.ToString().Trim(), 10, '0');
                        C068C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C068C").Value.ToString().Trim(), 10, '0');
                        C069 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C069").Value.ToString().Trim(), 10, '0');
                        C070 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C070").Value.ToString().Trim(), 10, '0');

                        C071 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C071").Value.ToString().Trim(), 10, '0');
                        C072 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C072").Value.ToString().Trim(), 10, '0');
                        C073 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C073").Value.ToString().Trim(), 10, '0');
                        C074 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C074").Value.ToString().Trim(), 10, '0');
                        C075 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C075").Value.ToString().Trim(), 10, '0');
                        C076 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C076").Value.ToString().Trim(), 11, '0');
                        C077 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C077").Value.ToString().Trim(), 10, '0');
                        C078 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C078").Value.ToString().Trim(), 11, '0');
                        C079 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C079").Value.ToString().Trim(), 8, '0');
                        C080 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C080").Value.ToString().Trim(), 8, '0');

                        C081A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C081A").Value.ToString().Trim(), 2, '0');
                        C081B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C081B").Value.ToString().Trim(), 8, '0');
                        C082A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C082A").Value.ToString().Trim(), 2, '0');
                        C082B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C082B").Value.ToString().Trim(), 8, '0');
                        C083A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C083A").Value.ToString().Trim(), 2, '0');
                        C083B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C083B").Value.ToString().Trim(), 8, '0');
                        C084 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C084").Value.ToString().Trim(), 8, '0');
                        C085 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C085").Value.ToString().Trim(), 10, '0');
                        C086A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C086A").Value.ToString().Trim(), 10, '0');
                        C086B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C086B").Value.ToString().Trim(), 10, '0');
                        C087A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C087A").Value.ToString().Trim(), 10, '0');
                        C087B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C087B").Value.ToString().Trim(), 10, '0');
                        C088A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C088A").Value.ToString().Trim(), 10, '0');
                        C088B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C088B").Value.ToString().Trim(), 10, '0');
                        C089A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C089A").Value.ToString().Trim(), 10, '0');
                        C089B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C089B").Value.ToString().Trim(), 10, '0');
                        C090A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090A").Value.ToString().Trim(), 10, '0');
                        C090B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C090B").Value.ToString().Trim(), 10, '0');

                        C091A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C091A").Value.ToString().Trim(), 10, '0');
                        C091B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C091B").Value.ToString().Trim(), 10, '0');
                        C092A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C092A").Value.ToString().Trim(), 10, '0');
                        C092B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C092B").Value.ToString().Trim(), 10, '0');
                        C093A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C093A").Value.ToString().Trim(), 8, '0');
                        C093B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C093B").Value.ToString().Trim(), 8, '0');
                        C094A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C094A").Value.ToString().Trim(), 8, '0');
                        C094B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C094B").Value.ToString().Trim(), 8, '0');
                        C094C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C094C").Value.ToString().Trim(), 8, '0');
                        C095A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C095A").Value.ToString().Trim(), 8, '0');
                        C095B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C095B").Value.ToString().Trim(), 8, '0');
                        C096A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C096A").Value.ToString().Trim(), 8, '0');
                        C096B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C096B").Value.ToString().Trim(), 8, '0');
                        C096C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C096C").Value.ToString().Trim(), 8, '0');
                        C096D = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C096D").Value.ToString().Trim(), 8, '0');
                        C097 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C097").Value.ToString().Trim(), 11, '0');
                        C098 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C098").Value.ToString().Trim(), 11, '0');
                        C099 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C099").Value.ToString().Trim(), 11, '0');
                        C100 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C100").Value.ToString().Trim(), 11, '0');

                        C101 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C101").Value.ToString().Trim(), 11, '0');
                        C102 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C102").Value.ToString().Trim(), 8, '0');
                        C103 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C103").Value.ToString().Trim(), 10, '0');
                        C104 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C104").Value.ToString().Trim(), 10, '0');
                        C105 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C105").Value.ToString().Trim(), 10, '0');
                        C106 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C106").Value.ToString().Trim(), 10, '0');
                        C107 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C107").Value.ToString().Trim(), 10, '0');
                        C108 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C108").Value.ToString().Trim(), 8, '0');
                        C109 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C109").Value.ToString().Trim(), 10, '0');
                        C110 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C110").Value.ToString().Trim(), 10, '0');

                        C111 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C111").Value.ToString().Trim(), 10, '0');
                        C112 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C112").Value.ToString().Trim(), 10, '0');
                        C113 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C113").Value.ToString().Trim(), 10, '0');
                        C114 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C114").Value.ToString().Trim(), 11, '0');
                        C115 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C115").Value.ToString().Trim(), 11, '0');
                        C116 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C116").Value.ToString().Trim(), 11, '0');
                        C117 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C117").Value.ToString().Trim(), 11, '0');
                        C118 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C118").Value.ToString().Trim(), 10, '0');
                        C119 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C119").Value.ToString().Trim(), 10, '0');
                        C120 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C120").Value.ToString().Trim(), 10, '0');

                        C121 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C121").Value.ToString().Trim(), 10, '0');
                        C122 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C122").Value.ToString().Trim(), 10, '0');
                        C123 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C123").Value.ToString().Trim(), 10, '0');
                        C124 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C124").Value.ToString().Trim(), 10, '0');
                        C125 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C125").Value.ToString().Trim(), 10, '0');
                        C126A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C126A").Value.ToString().Trim(), 2, '0');
                        C126B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C126B").Value.ToString().Trim(), 10, '0');
                        C127A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C127A").Value.ToString().Trim(), 2, '0');
                        C127B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C127B").Value.ToString().Trim(), 10, '0');
                        C128A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C128A").Value.ToString().Trim(), 10, '0');
                        C128B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C128B").Value.ToString().Trim(), 10, '0');
                        C129A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C129A").Value.ToString().Trim(), 10, '0');
                        C129B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C129B").Value.ToString().Trim(), 10, '0');
                        C130A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C130A").Value.ToString().Trim(), 10, '0');
                        C130B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C130B").Value.ToString().Trim(), 10, '0');

                        C131A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C131A").Value.ToString().Trim(), 10, '0');
                        C131B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C131B").Value.ToString().Trim(), 10, '0');
                        C132A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C132A").Value.ToString().Trim(), 10, '0');
                        C132B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C132B").Value.ToString().Trim(), 10, '0');
                        C133A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C133A").Value.ToString().Trim(), 10, '0');
                        C133B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C133B").Value.ToString().Trim(), 10, '0');
                        C134A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C134A").Value.ToString().Trim(), 10, '0');
                        C134B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C134B").Value.ToString().Trim(), 10, '0');
                        C135A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C135A").Value.ToString().Trim(), 10, '0');
                        C135B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C135B").Value.ToString().Trim(), 10, '0');
                        C136A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C136A").Value.ToString().Trim(), 11, '0');
                        C136B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C136B").Value.ToString().Trim(), 10, '0');
                        C137A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C137A").Value.ToString().Trim(), 11, '0');
                        C137B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C137B").Value.ToString().Trim(), 10, '0');
                        C138A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C138A").Value.ToString().Trim(), 11, '0');
                        C138B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C138B").Value.ToString().Trim(), 10, '0');
                        C139A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C139A").Value.ToString().Trim(), 11, '0');
                        C139B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C139B").Value.ToString().Trim(), 10, '0');
                        C140A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C140A").Value.ToString().Trim(), 11, '0');
                        C140B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C140B").Value.ToString().Trim(), 10, '0');

                        C141 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C141").Value.ToString().Trim(), 11, '0');
                        C142 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C142").Value.ToString().Trim(), 11, '0');
                        C143 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C143").Value.ToString().Trim(), 10, '0');
                        C144 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C144").Value.ToString().Trim(), 10, '0');
                        C145 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C145").Value.ToString().Trim(), 10, '0');
                        C146 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C146").Value.ToString().Trim(), 10, '0');
                        C147 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C147").Value.ToString().Trim(), 10, '0');
                        C148A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C148A").Value.ToString().Trim(), 10, '0');
                        C148B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C148B").Value.ToString().Trim(), 8, '0');
                        C149 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C149").Value.ToString().Trim(), 10, '0');
                        C150 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C150").Value.ToString().Trim(), 10, '0');

                        C151 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C151").Value.ToString().Trim(), 10, '0');
                        C152A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C152A").Value.ToString().Trim(), 11, '0');
                        C152B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C152B").Value.ToString().Trim(), 10, '0');
                        C152C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C152C").Value.ToString().Trim(), 10, '0');
                        C153 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C153").Value.ToString().Trim(), 3, '0');
                        C154A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C154A").Value.ToString().Trim(), 11, '0');
                        C154B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C154B").Value.ToString().Trim(), 10, '0');
                        C154C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C154C").Value.ToString().Trim(), 10, '0');
                        C155A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C155A").Value.ToString().Trim(), 11, '0');
                        C155B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C155B").Value.ToString().Trim(), 10, '0');
                        C155C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C155C").Value.ToString().Trim(), 10, '0');
                        C156A_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156A_1").Value.ToString().Trim(), 1, '0');
                        C156A_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156A_2").Value.ToString().Trim(), 11, '0');
                        C156B_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156B_1").Value.ToString().Trim(), 1, '0');
                        C156B_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156B_2").Value.ToString().Trim(), 10, '0');
                        C156C_1 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156C_1").Value.ToString().Trim(), 1, '0');
                        C156C_2 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C156C_2").Value.ToString().Trim(), 10, '0');

                        FileSystem.PrintLine(1, C001 + C002 + C003 + C004 + C005 + C006 + C007 + C008 + C009 + C010 + C011 + C012 + C013 + C014 + C015 + C016 + C017 + C018 + C019 + C020
                                              + C021 + C022 + C023 + C024 + C025 + C026 + C027 + C028 + C029 + C030 + C031 + C032 + C033 + C034 + C035 + C036 + C037 + C038 + C039 + C040
                                              + C041 + C042 + C043 + C044 + C045 + C046 + C047 + C048 + C049 + C050 + C051 + C052 + C053 + C054 + C055 + C056 + C057 + C058 + C059 + C060
                                              + C061 + C062A + C062B + C062C + C063 + C064 + C065 + C066 + C067 + C068A + C068B + C068C + C069 + C070 + C071 + C072 + C073 + C074 + C075 + C076 + C077 + C078 + C079 + C080
                                              + C081A + C081B + C082A + C082B + C083A + C083B + C084 + C085 + C086A + C086B + C087A + C087B + C088A + C088B + C089A + C089B + C090A + C090B + C091A + C091B + C092A + C092B + C093A + C093B + C094A + C094B + C094C + C095A + C095B + C096A + C096B + C096C + C096D + C097 + C098 + C099 + C100
                                              + C101 + C102 + C103 + C104 + C105 + C106 + C107 + C108 + C109 + C110 + C111 + C112 + C113 + C114 + C115 + C116 + C117 + C118 + C119 + C120
                                              + C121 + C122 + C123 + C124 + C125 + C126A + C126B + C127A + C127B + C128A + C128B + C129A + C129B + C130A + C130B + C131A + C131B + C132A + C132B + C133A + C133B + C134A + C134B + C135A + C135B + C136A + C136B + C137A + C137B + C138A + C138B + C139A + C139B + C140A + C140B
                                              + C141 + C142 + C143 + C144 + C145 + C146 + C147 + C148A + C148B + C149 + C150 + C151 + C152A + C152B + C152C + C153 + C154A + C154B + C154C + C155A + C155B + C155C + C156A_1 + C156A_2 + C156B_1 + C156B_2 + C156C_1 + C156C_2);

                        // D 레코드: 종전근무처 레코드
                        if (Conversion.Val(C006) > 0)
                        {
                            if (File_Create_D_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                            {
                                errNum = 2;
                                throw new Exception();
                            }
                        }

                        // E 레코드: 부양가족 레코드
                        if (File_Create_E_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                        {
                            errNum = 3;
                            throw new Exception();
                        }

                        // F 레코드: 연금.저축 등 소득.세액 공제명세 레코드
                        if (File_Create_F_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                        {
                            errNum = 4;
                            throw new Exception();
                        }

                        // G 레코드: 월세.거주자간 주택임차차임금 원리금 상환액 소득공제명세 레코드
                        if (File_Create_G_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                        {
                            errNum = 5;
                            throw new Exception();
                        }

                        // H 레코드: 기부조정명세 레코드
                        if (File_Create_H_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                        {
                            errNum = 6;
                            throw new Exception();
                        }

                        // I 레코드 : 해당년도 기부명세 레코드
                        if (File_Create_I_record(c_SAUP, c_YYYY, c_SABUN, C004) == false)
                        {
                            errNum = 7;
                            throw new Exception();
                        }

                        oRecordSet.MoveNext();

                        ProgressBar01.Value +=  1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";
                    }
                }

                returnValue = true;
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// D 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_D_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = false;
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
            string D055;  // 10   '비과세(T50:내국인우수인력국내복귀소득세감면)
            string D056A; // 10   '비과세(T11:중소기업취업청년소득세감면50%)
            string D056B; // 10   '비과세(T12:중소기업취업청년소득세감면70%)
            string D056C; // 10   '비과세(T13:중소기업취업청년소득세감면90%)
            string D057;  // 10   '비과세(T20:조세조약상교직자감면)
            string D058;  // 10   '공란  9(10)
            string D059;  // 10   '공란  9(10)
            string D060;  // 10   '공란  9(10)
            string D061;  // 10   '공란  9(10)
            string D062;  // 10   '비과세 계
            string D063;  // 10   '감면소득 계
            string D064A; // 11   '소득세
            string D064B; // 10   '지방소득세
            string D064C; // 10   '농특세
            string D065;  // 2    '종(전)근무처일련번호 
            string D066;  // 1172 '공란

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
                        D055 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D055").Value.ToString().Trim(), 10, '0');
                        D056A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D056A").Value.ToString().Trim(), 10, '0');
                        D056B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D056B").Value.ToString().Trim(), 10, '0');
                        D056C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D056C").Value.ToString().Trim(), 10, '0');
                        D057 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D057").Value.ToString().Trim(), 10, '0');
                        D058 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D058").Value.ToString().Trim(), 10, '0');
                        D059 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D059").Value.ToString().Trim(), 10, '0');
                        D060 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D060").Value.ToString().Trim(), 10, '0');
                        D061 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D061").Value.ToString().Trim(), 10, '0');
                        D062 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D062").Value.ToString().Trim(), 10, '0');
                        D063 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D063").Value.ToString().Trim(), 10, '0');
                        D064A = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D064A").Value.ToString().Trim(), 11, '0');
                        D064B = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D064B").Value.ToString().Trim(), 10, '0');
                        D064C = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D064C").Value.ToString().Trim(), 10, '0');
                        D065 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D065").Value.ToString().Trim(), 2);
                        D066 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("D066").Value.ToString().Trim(), 1172);

                        FileSystem.PrintLine(1, D001 + D002 + D003 + D004 + D005 + D006 + D007 + D008 + D009 + D010 + D011 + D012 + D013 + D014 + D015 + D016 + D017 + D018 + D019 + D020
                                              + D021 + D022 + D023 + D024 + D025 + D026 + D027 + D028 + D029 + D030 + D031 + D032 + D033 + D034 + D035 + D036 + D037 + D038 + D039 + D040
                                              + D041 + D042 + D043 + D044 + D045 + D046 + D047 + D048 + D049 + D050A + D050B + D050C + D051 + D052 + D053 + D054 + D055 + D056A + D056B + D056C + D057 + D058 + D059 + D060
                                              + D061 + D062 + D063 + D064A + D064B + D064C + D065 + D066);

                        oRecordSet.MoveNext();
                    }
                }

                returnValue = true;
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

            return returnValue;
        }

        /// <summary>
        /// E 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_E_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = false;
            short errNum = 0;
            int i = 0;
            int BUYCNT = 0;
            int FAMCNT = 0;
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
            // 2020 년 3개로
            string[] E007 = new string[3]; // 1    '관계코드
            string[] E008 = new string[3]; // 1    '내외국인구분
            string[] E009 = new string[3]; // 30   '성명
            string[] E010 = new string[3]; // 13   '주민등록번호
            string[] E011 = new string[3]; // 1    '기본공제
            string[] E012 = new string[3]; // 1    '장애자공제
            string[] E013 = new string[3]; // 1    '부녀자공제
            string[] E014 = new string[3]; // 1    '경로우대
            string[] E015 = new string[3]; // 1    '한부모공제
            string[] E016 = new string[3]; // 1    '출산.입양공제
            string[] E017 = new string[3]; // 1    '자녀공제
            string[] E018 = new string[3]; // 1    '교육비공제 1,2,3,4
            string[] E019 = new string[3]; // 10   '국세청-보험료_건강보험
            string[] E020 = new string[3]; // 10   '국세청-보험료_고용보험
            string[] E021 = new string[3]; // 10   '국세청-보험료_보장성보험
            string[] E022 = new string[3]; // 10   '국세청-보험료_장애인전용보장성보험
            string[] E023 = new string[3]; // 10   '국세청-의료비_일반
            string[] E024 = new string[3]; // 10   '국세청-의료비_난임
            string[] E025 = new string[3]; // 10   '국세청-의료비_65세이상.장애인.건강보험산정특례자
            string[] E026 = new string[3]; // 10   '국세청-의료비_실손의료보험금
            string[] E027 = new string[3]; // 10   '국세청-교육비_일반
            string[] E028 = new string[3]; // 10   '국세청-교육비_장애인특수교육비
            // 2020년 신용카드 등
            string[] E029 = new string[3]; // 10   '국세청-신용카드  3월
            string[] E030 = new string[3]; // 10   '국세청-직불카드  3월
            string[] E031 = new string[3]; // 10   '국세청-현금영수증  3월
            string[] E032 = new string[3]; // 10   '국세청-도서.공연사용분 3월
            string[] E033 = new string[3]; // 10   '국세청-전통시장사용액 3월
            string[] E034 = new string[3]; // 10   '국세청-대중교통이용액 3월
            string[] E035 = new string[3]; // 10   '국세청-신용카드  4-7월
            string[] E036 = new string[3]; // 10   '국세청-직불카드  4-7월
            string[] E037 = new string[3]; // 10   '국세청-현금영수증  4-7월
            string[] E038 = new string[3]; // 10   '국세청-도서.공연사용분 4-7월
            string[] E039 = new string[3]; // 10   '국세청-전통시장사용액 4-7월
            string[] E040 = new string[3]; // 10   '국세청-대중교통이용액 4-7월
            string[] E041 = new string[3]; // 10   '국세청-신용카드  그외
            string[] E042 = new string[3]; // 10   '국세청-직불카드  그외
            string[] E043 = new string[3]; // 10   '국세청-현금영수증  그외
            string[] E044 = new string[3]; // 10   '국세청-도서.공연사용분 그외
            string[] E045 = new string[3]; // 10   '국세청-전통시장사용액 그외
            string[] E046 = new string[3]; // 10   '국세청-대중교통이용액 그외
            //
            string[] E047 = new string[3]; // 13   '국세청-기부금

            string[] E048 = new string[3]; // 10   '기타-보험료_건강보험
            string[] E049 = new string[3]; // 10   '기타-보험료_고용보험
            string[] E050 = new string[3]; // 10   '기타-보험료_보장성
            string[] E051 = new string[3]; // 10   '기타-보험료_장애인전용보장성
            string[] E052 = new string[3]; // 10   '기타-의료비_일반
            string[] E053 = new string[3]; // 10   '기타-의료비_난임
            string[] E054 = new string[3]; // 10   '기타-의료비_65세이상.장애인.건강보험산정특례자
            string[] E055_1 = new string[3]; // 1  '기타-의료비_실손의료보험금부호
            string[] E055_2 = new string[3]; // 10 '기타-의료비_실손의료보험금
            string[] E056 = new string[3]; // 10   '기타-교육비_일반
            string[] E057 = new string[3]; // 10   '기타-교육비_장애인특수교육비
            // 2020년 신용카드 등
            string[] E058 = new string[3]; // 10   '기타-신용카드   3월
            string[] E059 = new string[3]; // 10   '기타-직불카드   3월
            string[] E060 = new string[3]; // 10   '기타-도서.공연사용분 3월
            string[] E061 = new string[3]; // 10   '기타-전통시장사용액  3월
            string[] E062 = new string[3]; // 10   '기타-대중교통이용액  3월 
            string[] E063 = new string[3]; // 10   '기타-신용카드   3월
            string[] E064 = new string[3]; // 10   '기타-직불카드   3월
            string[] E065 = new string[3]; // 10   '기타-도서.공연사용분 3월
            string[] E066 = new string[3]; // 10   '기타-전통시장사용액  3월
            string[] E067 = new string[3]; // 10   '기타-대중교통이용액  3월 
            string[] E068 = new string[3]; // 10   '기타-신용카드   3월
            string[] E069 = new string[3]; // 10   '기타-직불카드   3월
            string[] E070 = new string[3]; // 10   '기타-도서.공연사용분 3월
            string[] E071 = new string[3]; // 10   '기타-전통시장사용액  3월
            string[] E072 = new string[3]; // 10   '기타-대중교통이용액  3월 
            //
            string[] E073 = new string[3]; // 13   '기타-기부금

            string E208;                   // 2    '부양가족레코드일련번호
            string E209 = string.Empty;    // 26   '공란

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
                        if (BUYCNT == 0)  // 초기화
                        {
                            for (i = 0; i <= 2; i++)   //  2020 3개
                            {
                                E007[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E008[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E009[i] = codeHelpClass.GetFixedLengthStringByte("", 30);
                                E010[i] = codeHelpClass.GetFixedLengthStringByte("", 13);
                                E011[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E012[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E013[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E014[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E015[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E016[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E017[i] = codeHelpClass.GetFixedLengthStringByte("", 1);
                                E018[i] = codeHelpClass.GetFixedLengthStringByte("", 1);

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
                                E036[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
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
                                E047[i] = codeHelpClass.GetFixedLengthStringByte("0", 13, '0');  // 13
                                E048[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E049[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E050[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E051[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E052[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E053[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E054[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E055_1[i] = codeHelpClass.GetFixedLengthStringByte("0", 1, '0');
                                E055_2[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E056[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E057[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E058[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E059[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E060[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E061[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E062[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E063[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E064[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E065[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E066[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E067[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E068[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E069[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E070[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E071[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E072[i] = codeHelpClass.GetFixedLengthStringByte("0", 10, '0');
                                E073[i] = codeHelpClass.GetFixedLengthStringByte("0", 13, '0');  // 13
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
                        E036[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E036").Value.ToString().Trim(), 10, '0');
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
                        E047[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E047").Value.ToString().Trim(), 13, '0'); // 13
                        E048[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E048").Value.ToString().Trim(), 10, '0');
                        E049[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E049").Value.ToString().Trim(), 10, '0');
                        E050[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E050").Value.ToString().Trim(), 10, '0');
                        E051[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E051").Value.ToString().Trim(), 10, '0');
                        E052[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E052").Value.ToString().Trim(), 10, '0');
                        E053[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E053").Value.ToString().Trim(), 10, '0');
                        E054[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E054").Value.ToString().Trim(), 10, '0');
                        E055_1[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E055_1").Value.ToString().Trim(), 1, '0');
                        E055_2[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E055_2").Value.ToString().Trim(), 10, '0');
                        E056[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E056").Value.ToString().Trim(), 10, '0');
                        E057[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E057").Value.ToString().Trim(), 10, '0');
                        E058[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E058").Value.ToString().Trim(), 10, '0');
                        E059[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E059").Value.ToString().Trim(), 10, '0');
                        E060[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E060").Value.ToString().Trim(), 10, '0');
                        E061[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E061").Value.ToString().Trim(), 10, '0');
                        E062[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E062").Value.ToString().Trim(), 10, '0');
                        E063[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E063").Value.ToString().Trim(), 10, '0');
                        E064[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E064").Value.ToString().Trim(), 10, '0');
                        E065[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E065").Value.ToString().Trim(), 10, '0');
                        E066[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E066").Value.ToString().Trim(), 10, '0');
                        E067[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E067").Value.ToString().Trim(), 10, '0');
                        E068[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E068").Value.ToString().Trim(), 10, '0');
                        E069[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E069").Value.ToString().Trim(), 10, '0');
                        E070[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E070").Value.ToString().Trim(), 10, '0');
                        E071[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E071").Value.ToString().Trim(), 10, '0');
                        E072[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E072").Value.ToString().Trim(), 10, '0');
                        E073[BUYCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E073").Value.ToString().Trim(), 13, '0'); // 13

                        E209 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("E209").Value.ToString().Trim(), 26); //공란

                        oRecordSet.MoveNext();

                        // If BUYCNT = 4 Then    '5개면 인쇄 0 - 4
                        if (BUYCNT == 2 || oRecordSet.EoF)  // 2020년 3개
                        {
                            E208 = codeHelpClass.GetFixedLengthStringByte(FAMCNT.ToString(), 2, '0'); // 일련번호

                            // E레코드 WRITE
                            FileSystem.PrintLine(1, E001 + E002 + E003 + E004 + E005 + E006
                                                  + E007[0] + E008[0] + E009[0] + E010[0] + E011[0] + E012[0] + E013[0] + E014[0] + E015[0] + E016[0] + E017[0] + E018[0] + E019[0] + E020[0]
                                                  + E021[0] + E022[0] + E023[0] + E024[0] + E025[0] + E026[0] + E027[0] + E028[0] + E029[0] + E030[0] + E031[0] + E032[0] + E033[0] + E034[0]
                                                  + E035[0] + E036[0] + E037[0] + E038[0] + E039[0] + E040[0] + E041[0] + E042[0] + E043[0] + E044[0] + E045[0] + E046[0] + E047[0] + E048[0]
                                                  + E049[0] + E050[0] + E051[0] + E052[0] + E053[0] + E054[0] + E055_1[0] + E055_2[0] + E056[0] + E057[0] + E058[0] + E059[0] + E060[0] + E061[0] + E062[0]
                                                  + E063[0] + E064[0] + E065[0] + E066[0] + E067[0] + E068[0] + E069[0] + E070[0] + E071[0] + E072[0] + E073[0]
                                                  + E007[1] + E008[1] + E009[1] + E010[1] + E011[1] + E012[1] + E013[1] + E014[1] + E015[1] + E016[1] + E017[1] + E018[1] + E019[1] + E020[1]
                                                  + E021[1] + E022[1] + E023[1] + E024[1] + E025[1] + E026[1] + E027[1] + E028[1] + E029[1] + E030[1] + E031[1] + E032[1] + E033[1] + E034[1]
                                                  + E035[1] + E036[1] + E037[1] + E038[1] + E039[1] + E040[1] + E041[1] + E042[1] + E043[1] + E044[1] + E045[1] + E046[1] + E047[1] + E048[1]
                                                  + E049[1] + E050[1] + E051[1] + E052[1] + E053[1] + E054[1] + E055_1[1] + E055_2[1] + E056[1] + E057[1] + E058[1] + E059[1] + E060[1] + E061[1] + E062[1]
                                                  + E063[1] + E064[1] + E065[1] + E066[1] + E067[1] + E068[1] + E069[1] + E070[1] + E071[1] + E072[1] + E073[1]
                                                  + E007[2] + E008[2] + E009[2] + E010[2] + E011[2] + E012[2] + E013[2] + E014[2] + E015[2] + E016[2] + E017[2] + E018[2] + E019[2] + E020[2]
                                                  + E021[2] + E022[2] + E023[2] + E024[2] + E025[2] + E026[2] + E027[2] + E028[2] + E029[2] + E030[2] + E031[2] + E032[2] + E033[2] + E034[2]
                                                  + E035[2] + E036[2] + E037[2] + E038[2] + E039[2] + E040[2] + E041[2] + E042[2] + E043[2] + E044[2] + E045[2] + E046[2] + E047[2] + E048[2]
                                                  + E049[2] + E050[2] + E051[2] + E052[2] + E053[2] + E054[2] + E055_1[2] + E055_2[2] + E056[2] + E057[2] + E058[2] + E059[2] + E060[2] + E061[2] + E062[2]
                                                  + E063[2] + E064[2] + E065[2] + E066[2] + E067[2] + E068[2] + E069[2] + E070[2] + E071[2] + E072[2] + E073[2]
                                                  //+ E007[3] + E008[3] + E009[3] + E010[3] + E011[3] + E012[3] + E013[3] + E014[3] + E015[3] + E016[3] + E017[3] + E018[3] + E019[3] + E020[3]
                                                  //+ E021[3] + E022[3] + E023[3] + E024[3] + E025[3] + E026[3] + E027[3] + E028[3] + E029[3] + E030[3] + E031[3] + E032[3] + E033[3] + E034[3]
                                                  //+ E035[3] + E036[3] + E037[3] + E038[3] + E039[3] + E040[3] + E041[3] + E042[3] + E043[3] + E044[3] + E045[3] + E046[3] + E047[3] + E048[3]
                                                  //+ E049[3] + E050[3] + E051[3] + E052[3] + E053[3] + E054[3] + E055[3] + E056[3] + E057[3] + E058[3] + E059[3] + E060[3] + E061[3] + E062[3]
                                                  //+ E063[3] + E064[3] + E065[3] + E066[3] + E067[3] + E068[3] + E069[3] + E070[3] + E071[3] + E072[3] + E073[3]
                                                  //+ E007[4] + E008[4] + E009[4] + E010[4] + E011[4] + E012[4] + E013[4] + E014[4] + E015[4] + E016[4] + E017[4] + E018[4] + E019[4] + E020[4]
                                                  //+ E021[4] + E022[4] + E023[4] + E024[4] + E025[4] + E026[4] + E027[4] + E028[4] + E029[4] + E030[4] + E031[4] + E032[4] + E033[4] + E034[4]
                                                  //+ E035[4] + E036[4] + E037[4] + E038[4] + E039[4] + E040[4] + E041[4] + E042[4] + E043[4] + E044[4] + E045[4] + E046[4] + E047[4] + E048[4]
                                                  //+ E049[4] + E050[4] + E051[4] + E052[4] + E053[4] + E054[4] + E055[4] + E056[4] + E057[4] + E058[4] + E059[4] + E060[4] + E061[4] + E062[4]
                                                  //+ E063[4] + E064[4] + E065[4] + E066[4] + E067[4] + E068[4] + E069[4] + E070[4] + E071[4] + E072[4] + E073[4]
                                                  + E208 + E209);
                            BUYCNT = 0;
                            FAMCNT += 1;
                        }
                        else
                        {
                            BUYCNT += 1; // 해당사원의 부양가족일련번호 +1 하기
                        }
                    }
                }
                else
                {
                    errNum = 1;
                    throw new Exception();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
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

            return returnValue;
        }

        /// <summary>
        /// F 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_F_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = true;  // 기본을 TRUE 로
            int i = 0;
            int sCNT = 0;
            int rCNT = 0;
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

            string F127;                    // 2    '연금.저축레코드일련번호
            string F128 = string.Empty;     // 206  '공란

            try
            {
                rCNT = 1;
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

                    sCNT = 0;
                    while (!oRecordSet.EoF)
                    {
                        // 초기화
                        if (sCNT == 0)
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

                        F007[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F007").Value.ToString().Trim(), 2);
                        F008[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F008").Value.ToString().Trim(), 3);
                        F009[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F009").Value.ToString().Trim(), 60);
                        F010[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F010").Value.ToString().Trim(), 20);
                        F011[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F011").Value.ToString().Trim(), 10, '0');
                        F012[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F012").Value.ToString().Trim(), 10, '0');
                        F013[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F013").Value.ToString().Trim(), 4, '0');
                        F014[sCNT] = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F014").Value.ToString().Trim(), 1);

                        F128 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("F128").Value.ToString().Trim(), 206);

                        oRecordSet.MoveNext();

                        // If sCNT 가 15개나 끝이면 인쇄
                        if (sCNT == 14 | oRecordSet.EoF)
                        {
                            F127 = codeHelpClass.GetFixedLengthStringByte(rCNT.ToString(), 2, '0'); // 일련번호
                            // F 레코드 WRITE
                            FileSystem.PrintLine(1, F001 + F002 + F003 + F004 + F005 + F006
                                                  + F007[0] + F008[0] + F009[0] + F010[0] + F011[0] + F012[0] + F013[0] + F014[0] + F007[1] + F008[1] + F009[1] + F010[1] + F011[1] + F012[1] + F013[1] + F014[1]
                                                  + F007[2] + F008[2] + F009[2] + F010[2] + F011[2] + F012[2] + F013[2] + F014[2] + F007[3] + F008[3] + F009[3] + F010[3] + F011[3] + F012[3] + F013[3] + F014[3]
                                                  + F007[4] + F008[4] + F009[4] + F010[4] + F011[4] + F012[4] + F013[4] + F014[4] + F007[5] + F008[5] + F009[5] + F010[5] + F011[5] + F012[5] + F013[5] + F014[5]
                                                  + F007[6] + F008[6] + F009[6] + F010[6] + F011[6] + F012[6] + F013[6] + F014[6] + F007[7] + F008[7] + F009[7] + F010[7] + F011[7] + F012[7] + F013[7] + F014[7]
                                                  + F007[8] + F008[8] + F009[8] + F010[8] + F011[8] + F012[8] + F013[8] + F014[8] + F007[9] + F008[9] + F009[9] + F010[9] + F011[9] + F012[9] + F013[9] + F014[9]
                                                  + F007[10] + F008[10] + F009[10] + F010[10] + F011[10] + F012[10] + F013[10] + F014[10] + F007[11] + F008[11] + F009[11] + F010[11] + F011[11] + F012[11] + F013[11] + F014[11]
                                                  + F007[12] + F008[12] + F009[12] + F010[12] + F011[12] + F012[12] + F013[12] + F014[12] + F007[13] + F008[13] + F009[13] + F010[13] + F011[13] + F012[13] + F013[13] + F014[13]
                                                  + F007[14] + F008[14] + F009[14] + F010[14] + F011[14] + F012[14] + F013[14] + F014[14] + F127 + F128);
                            sCNT = 0;
                            rCNT += 1;
                        }
                        else
                        {
                            sCNT += 1; // 레코드번호 + 1
                        }
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                returnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// G 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_G_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = true;  // 기본을 TRUE 로
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
            string G086;  // 197  '공란

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
                    G086 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("G086").Value.ToString().Trim(), 197);

                    // G 레코드 삽입
                    FileSystem.PrintLine(1, G001 + G002 + G003 + G004 + G005 + G006 + G007 + G008 + G009 + G010 + G011 + G012 + G013 + G014 + G015 + G016 + G017 + G018 + G019 + G020
                                          + G021 + G022 + G023 + G024 + G025 + G026 + G027 + G028 + G029 + G030 + G031 + G032 + G033 + G034 + G035 + G036 + G037 + G038 + G039 + G040
                                          + G041 + G042 + G043 + G044 + G045 + G046 + G047 + G048 + G049 + G050 + G051 + G052 + G053 + G054 + G055 + G056 + G057 + G058 + G059 + G060
                                          + G061 + G062 + G063 + G064 + G065 + G066 + G067 + G068 + G069 + G070 + G071 + G072 + G073 + G074 + G075 + G076 + G077 + G078 + G079 + G080
                                          + G081 + G082 + G083 + G084 + G085 + G086);
                }
            }
            catch (Exception ex)
            {
                returnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// H 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_H_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = true;  // 기본을 TRUE 로
            int hCnt = 0;
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
            string H019;  // 1725  '공란

            try
            {
                // H_RECORE QUERY
                sQry = "EXEC PH_PY980_H '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    hCnt = 0;
                    while (!oRecordSet.EoF)
                    {
                        hCnt += 1; // 일련번호
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
                        H018 = codeHelpClass.GetFixedLengthStringByte(hCnt.ToString().Trim(), 5, '0');
                        H019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("H019").Value.ToString().Trim(), 1725);

                        // H 레코드 WRITE
                        FileSystem.PrintLine(1, H001 + H002 + H003 + H004 + H005 + H006 + H007 + H008 + H009 + H010
                                              + H011 + H012 + H013 + H014 + H015 + H016 + H017 + H018 + H019);
                        oRecordSet.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                returnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// I 레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_I_record(string psaup, string pyyyy, string psabun, string pC004)
        {
            bool returnValue = true;  // 기본을 TRUE 로
            int iCnt = 0;
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
            string I021;  // 1675 '공란

            try
            {
                // H_RECORE QUERY
                sQry = "EXEC PH_PY980_I '" + psaup + "', '" + pyyyy + "', '" + psabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    iCnt = 0;
                    while (!oRecordSet.EoF)
                    {
                        iCnt += 1; // 일련번호
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
                        I020 = codeHelpClass.GetFixedLengthStringByte(iCnt.ToString().Trim(), 5, '0');
                        I021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("I021").Value.ToString().Trim(), 1675);

                        // I 레코드 삽입
                        FileSystem.PrintLine(1, I001 + I002 + I003 + I004 + I005 + I006 + I007 + I008 + I009 + I010
                                              + I011 + I012 + I013 + I014 + I015 + I016 + I017 + I018 + I019 + I020 + I021);
                        oRecordSet.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                returnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                FileSystem.FileClose(1);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 필수 입력값 체크
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool returnValue = false;
            short errNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("HtaxID").Specific.Value))
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TeamName").Specific.Value))
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Dname").Specific.Value))
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Dtel").Specific.Value))
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errNum = 5;
                    throw new Exception();
                }

                returnValue = true;
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
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    break;
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
                        case "1288":
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
