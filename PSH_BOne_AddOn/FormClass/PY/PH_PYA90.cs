using Microsoft.VisualBasic;
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 근로소득간이지급명세서(세무서신고파일생성)
    /// </summary>
    internal class PH_PYA90 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private string CLTCOD;
        private string yyyy;
        private string Gubun;
        private string HtaxID;
        private string TeamName;
        private string Dname;
        private string Dtel;
        private string DocDate;

        /// <summary>
        /// Form 호출
        /// </summary>
        public override void LoadForm()
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PYA90.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PYA90_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PYA90");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
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
            int MM;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                MM = DateTime.Now.Month;

                // 귀속년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.Value = DateTime.Now.Year;

                // 근무기간
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                if (MM > 6 )
                {
                    oForm.Items.Item("Gubun").Specific.Value = "2";  // 하반기
                }
                else
                {
                    oForm.Items.Item("Gubun").Specific.Value = "1";  // 상반기
                }
                
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
            string stringSpace;

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
                Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
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
                if (File_Create_A_record() == false)
                {
                    errNum = 2;
                    throw new Exception();
                }

                // B RECORD 처리
                if (File_Create_B_record() == false)
                {
                    errNum = 3;
                    throw new Exception();
                }

                // C RECORD 처리
                if (File_Create_C_record() == false)
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
        private bool File_Create_A_record()
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry;
            string saup;
            string oFilePath; //파일 경로

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // A 제출자 레코드
            // 2019년귀속 190 BYTE

            string A001; // 1     '레코드구분
            string A002; // 2     '자료구분
            string A003; // 3     '세무서코드
            string A004; // 8     '제출일자
            string A005; // 1     '제출자구분 (1;세무대리인, 2;법인, 3;개인)
            string A006; // 6     '세무대리인
            string A007; // 20    '홈텍스ID
            string A008; // 4     '세무프로그램코드
            string A009; // 10    '사업자번호
            string A010; // 30    '법인명(상호)
            string A011; // 30    '담당자부서
            string A012; // 30    '담당자성명
            string A013; // 15    '담당자전화번호
            string A014; // 5     '신고의무자수
            string A015; // 25    '공란

            try
            {
                //A_RECORE QUERY
                sQry = "EXEC PH_PYA90_A '" + CLTCOD + "', '" + HtaxID + "', '" + TeamName + "', '" + Dname + "', '" + Dtel + "', '" + DocDate + "'";
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
                    oFilePath = "C:\\BANK\\SC" + codeHelpClass.Mid(saup, 0, 7) + "." + codeHelpClass.Mid(saup, 7, 3);
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
                    A010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A010").Value.ToString().Trim(), 30);
                    A011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A011").Value.ToString().Trim(), 30);
                    A012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A012").Value.ToString().Trim(), 30);
                    A013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A013").Value.ToString().Trim(), 15);
                    A014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A014").Value.ToString().Trim(), 5, '0');
                    A015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A015").Value.ToString().Trim(), 25);

                    FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015);

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
        private bool File_Create_B_record()
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // B 원천징수의무자별 집계 레코드
            string B001; // 1     '레코드구분
            string B002; // 2     '자료구분
            string B003; // 3     '세무서
            string B004; // 6     '일련번호
            string B005; // 40    '상호(법인명)
            string B006; // 30    '성명(대표자)
            string B007; // 10    '사업자번호
            string B008; // 13    '주민(법인)번호
            string B009; // 4     '귀속년도
            string B010; // 1     '근무시기  1.상반기 2.하반기
            string B011; // 10    '근로자수
            string B012; // 13    '과세소득합계
            string B013; // 13    '비과세소득합계
            string B014; // 44    '공란

            try
            {
                // B_RECORE QUERY
                sQry = "EXEC PH_PYA90_B '" + CLTCOD + "', '" + yyyy + "', '" + Gubun + "'";
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
                    B005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B005").Value.ToString().Trim(), 40);
                    B006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B006").Value.ToString().Trim(), 30);
                    B007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B007").Value.ToString().Trim(), 10);
                    B008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B008").Value.ToString().Trim(), 13);
                    B009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B009").Value.ToString().Trim(), 4);
                    B010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B010").Value.ToString().Trim(), 1);
                    B011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B011").Value.ToString().Trim(), 10, '0');
                    B012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B012").Value.ToString().Trim(), 13, '0');
                    B013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B013").Value.ToString().Trim(), 13, '0');
                    B014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("B014").Value.ToString().Trim(), 44);

                    FileSystem.PrintLine(1, B001 + B002 + B003 + B004 + B005 + B006 + B007 + B008 + B009 + B010 + B011 + B012 + B013 + B014);
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
        private bool File_Create_C_record()
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry;
            int NEWCNT; //일련번호

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            // C 주(현)근무지 레코드
            string C001;    // 1     '레코드구분
            string C002;    // 2     '자료구분
            string C003;    // 3     '세무서
            string C004;    // 7     '일련번호
            string C005;    // 10    '사업자번호
            string C006;    // 13    '주민등록번호
            string C007;    // 30    '성명
            string C008;    // 20    '전화번호
            string C009;    // 1     '내.외국인 1.내국인 9.외국인
            string C010;    // 1     '거주자구분 1.거주자 2.비거주자
            string C011;    // 2     '거주지국코드
            string C012;    // 8     '근무기간 시작연월일
            string C013;    // 8     '근무기간 종료연월일
            string C014;    // 13    '과세소득
            string C015;    // 13    '비과세소득
            string C016;    // 58    '공란

            try
            {
                // C_RECORE QUERY
                sQry = "EXEC PH_PYA90_C '" + CLTCOD + "', '" + yyyy + "', '" + Gubun + "'";
                oRecordSet.DoQuery(sQry);

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
                        NEWCNT += 1; //일련번호

                        //C RECORD MOVE
                        C001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C001").Value.ToString().Trim(), 1);
                        C002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C002").Value.ToString().Trim(), 2);
                        C003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C003").Value.ToString().Trim(), 3);
                        C004 = codeHelpClass.GetFixedLengthStringByte(NEWCNT.ToString(), 7, '0');
                        C005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C005").Value.ToString().Trim(), 10);
                        C006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C006").Value.ToString().Trim(), 13);
                        C007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C007").Value.ToString().Trim(), 30);
                        C008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C008").Value.ToString().Trim(), 20);
                        C009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C009").Value.ToString().Trim(), 1);
                        C010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C010").Value.ToString().Trim(), 1);
                        C011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C011").Value.ToString().Trim(), 2);
                        C012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C012").Value.ToString().Trim(), 8);
                        C013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C013").Value.ToString().Trim(), 8);
                        C014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C014").Value.ToString().Trim(), 13, '0');
                        C015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C015").Value.ToString().Trim(), 13, '0');
                        C016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("C016").Value.ToString().Trim(), 58);

                        FileSystem.PrintLine(1, C001 + C002 + C003 + C004 + C005 + C006 + C007 + C008 + C009 + C010 + C011 + C012 + C013 + C014 + C015 + C016);

                        oRecordSet.MoveNext();

                        ProgressBar01.Value += 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중...!";
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
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                FileSystem.FileClose(1);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }

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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
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
    }
}
