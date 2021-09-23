using System;
using SAPbouiCOM;
using System.Net;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.Code;
using System.IO;

namespace PSH_BOne_AddOn.Data
{
    /// <summary>
    /// SAP B1 Data 관련 Helper Class(SetMod, PS_Common 모듈의 내용 구현)
    /// </summary>
    public class PSH_DataHelpClass
    {
        /// <summary>
        /// DB에서 특정 필드값 조회 #1 (조건 추가)
        /// </summary>
        /// <param name="pReColumn">조회할 필드명</param>
        /// <param name="pColumn">조건절 필드명</param>
        /// <param name="pTable">테이블 명</param>
        /// <param name="pTaValue">조건절 비교문</param>
        /// <param name="pAndLine">조건절 추가 라인</param>
        /// <returns>pReColumn 필드의 내용</returns>
        public string Get_ReData(string pReColumn, string pColumn, string pTable, string pTaValue, string pAndLine)
        {
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string functionReturnValue = string.Empty;
            string sQry;

            sQry = "  SELECT " + pReColumn + " ";
            sQry += " FROM " + pTable;
            sQry += " WHERE " + pColumn + " = " + pTaValue;

            try
            {
                if (!string.IsNullOrEmpty(pAndLine))
                {
                    sQry += pAndLine;
                }

                oRecordSet.DoQuery(sQry);
                
                while (!oRecordSet.EoF)
                {
                    functionReturnValue = oRecordSet.Fields.Item(0).Value.ToString();
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// CHOOSEFROMLIST의 값을 리턴(임시주석처리 2019.03.18 송명규), 주석해제하여 구현중(2019.05.03 이후, 송명규)
        /// </summary>
        /// <param name="pVal">pVal</param>
        /// <param name="PSH_pFormUID">Fomr UID</param>
        /// <param name="PSH_pTableName">테이블이름</param>
        /// <param name="PSH_sUDS">리턴할 컬럼</param>
        /// <param name="PSH_pMatrix">Matrix</param>
        /// <param name="PSH_pRow">시작 Row</param>
        /// <param name="PSH_pSeqNoUDS">라인번호컬럼</param>
        /// <param name="PSH_pFieldName">체크박스일경우 컬럼명</param>
        /// <param name="PSH_pFieldValue">체크박스 초기값</param>
        public void PSH_CF_DBDatasourceReturn(SAPbouiCOM.ItemEvent pVal, string PSH_pFormUID, string PSH_pTableName, string PSH_sUDS, string PSH_pMatrix, short PSH_pRow, string PSH_pSeqNoUDS, string PSH_pFieldName, string PSH_pFieldValue)
        {
            SAPbouiCOM.IChooseFromListEvent PSH_oCFLEvento;
            SAPbouiCOM.ChooseFromList PSH_oCFL;
            SAPbouiCOM.DataTable PSH_oDataTable;
            SAPbouiCOM.Form PSH_pForm;
            SAPbouiCOM.Matrix PSH_oMatrix = null;
            SAPbouiCOM.DBDataSource PSH_oDBTable;

            short PSH_iLooper;
            short PSH_jLooper;
            string PSH_sCFLID;
            string[] PSH_sTemp01;
            
            PSH_pForm = PSH_Globals.SBO_Application.Forms.Item(PSH_pFormUID);
            PSH_oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
            PSH_oDataTable = PSH_oCFLEvento.SelectedObjects;
            PSH_sCFLID = PSH_oCFLEvento.ChooseFromListUID;
            // 취소버튼 클릭시
            if (PSH_oDataTable == null)
            {
                return;
            }

            PSH_oCFL = PSH_pForm.ChooseFromLists.Item(PSH_sCFLID);
            PSH_oDBTable = PSH_pForm.DataSources.DBDataSources.Item(PSH_pTableName);

            if (!string.IsNullOrEmpty(PSH_pMatrix))
            {
                PSH_oMatrix = PSH_pForm.Items.Item(PSH_pMatrix).Specific;
            }

            PSH_sTemp01 = PSH_sUDS.Split(','); //리턴할 컬럼의 이름을 배열에 저장

            if (!string.IsNullOrEmpty(PSH_pMatrix) && PSH_pRow > 0)
            {
                for (PSH_jLooper = 0; PSH_jLooper < PSH_oDataTable.Rows.Count; PSH_jLooper++)
                {
                    if (PSH_jLooper > 0)
                    {
                        if (!string.IsNullOrEmpty(PSH_pSeqNoUDS))
                        {
                            PSH_oDBTable.InsertRecord((PSH_pRow + PSH_jLooper - 1));
                            PSH_oDBTable.Offset = PSH_pRow + PSH_jLooper - 1;
                            PSH_oDBTable.SetValue(PSH_pSeqNoUDS, PSH_pRow + PSH_jLooper - 1, Convert.ToString(PSH_pRow + PSH_jLooper));
                        }
                        else
                        {
                            PSH_oDBTable.InsertRecord((PSH_pRow + PSH_jLooper - 1));
                            PSH_oDBTable.Offset = PSH_pRow + PSH_jLooper - 1;
                        }
                    }
                    else
                    {
                        PSH_oDBTable.Offset = PSH_pRow + PSH_jLooper - 1;
                    }

                    for (PSH_iLooper = 0; PSH_iLooper < PSH_sTemp01.Length; PSH_iLooper++)
                    {
                        // 사원마스타일경우 성 + 이름
                        if (PSH_oCFL.ObjectType == "171")
                        {
                            if (PSH_iLooper == 0)
                            {
                                PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], PSH_pRow + PSH_jLooper - 1, PSH_oDataTable.GetValue("U_MSTCOD", PSH_jLooper));
                            }
                            else if (PSH_iLooper == 1)
                            {
                                PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], PSH_pRow + PSH_jLooper - 1, PSH_oDataTable.GetValue("U_FULLNAME", PSH_jLooper));
                            }
                            else if (PSH_iLooper == 2)
                            {
                                PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], PSH_pRow + PSH_jLooper - 1, PSH_oDataTable.GetValue("U_TeamCode", PSH_jLooper));
                            }
                            else if (PSH_iLooper == 3)
                            {
                                PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], PSH_pRow + PSH_jLooper - 1, PSH_oDataTable.GetValue("U_TeamCode", PSH_jLooper));
                            }
                        }
                        else
                        {
                            PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], PSH_pRow + PSH_jLooper - 1, PSH_oDataTable.GetValue(PSH_iLooper, PSH_jLooper));
                        }
                    }

                    if (!string.IsNullOrEmpty(PSH_pFieldName) && !string.IsNullOrEmpty(PSH_pFieldValue))
                    {
                        PSH_oDBTable.SetValue(PSH_pFieldName, PSH_pRow + PSH_jLooper - 1, PSH_pFieldValue);
                    }

                    PSH_oMatrix.LoadFromDataSource();
                }
            }
            else
            {
                //PSH_sTemp02 = "";
                for (PSH_jLooper = 0; PSH_jLooper < PSH_oDataTable.Rows.Count; PSH_jLooper++)
                {
                    for (PSH_iLooper = 0; PSH_iLooper < PSH_sTemp01.Length; PSH_iLooper++)
                    {
                        switch (PSH_oCFL.ObjectType)
                        {
                            case "171":
                                break;

                            default:
                                PSH_oDBTable.SetValue(PSH_sTemp01[PSH_iLooper], 0, PSH_oDataTable.GetValue(PSH_iLooper, PSH_jLooper));
                                break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 일자(년월일) Validation 체크
        /// </summary>
        /// <param name="YearMon"></param>
        /// <returns></returns>
        public bool ChkYearMonth(string YearMon)
        {
            bool functionReturnValue;
            string oYear;
            string oMonth;

            if (YearMon.Length < 6)
            {
                functionReturnValue = false;
                return functionReturnValue;
            }

            oYear = YearMon.Substring(0, 4); //Strings.Mid(YearMon, 1, 4);

            if (Convert.ToInt16(oYear) < 2000 || Convert.ToInt16(oYear) > 3000)
            {
                functionReturnValue = false;
                return functionReturnValue;
            }

            oMonth = YearMon.Substring(4, 2); //Strings.Mid(YearMon, 5, 2);

            if (Convert.ToInt16(oMonth) < 1 || Convert.ToInt16(oMonth) > 12)
            {
                functionReturnValue = false;
                return functionReturnValue;
            }
            functionReturnValue = true;
            return functionReturnValue;
        }

        /// <summary>
        /// ComboBox 데이터 채우기
        /// </summary>
        /// <param name="pForm">화면</param>
        /// <param name="pSQL">쿼리</param>
        /// <param name="pCombo">콤보박스</param>
        /// <param name="pAddSpace">빈 값 추가 여부</param>
        public void SetReDataCombo(SAPbouiCOM.Form pForm, string pSQL, SAPbouiCOM.ComboBox pCombo, string pAddSpace)
        {
            int loopCount;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //기존 콤보 데이터 삭제
                if (pCombo.ValidValues.Count > 0)
                {
                    for (loopCount = 0; loopCount <= pCombo.ValidValues.Count - 1; loopCount++)
                    {
                        pCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }

                if (pAddSpace == "Y")
                {
                    pCombo.ValidValues.Add("", "");
                }

                oRecordSet.DoQuery(pSQL);

                if (oRecordSet.RecordCount > 0)
                {
                    for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
                    {
                        pCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// Matrix 콤보박스 세팅
        /// </summary>
        /// <param name="fCombo"></param>
        /// <param name="fSQL"></param>
        /// <param name="AndLine"></param>
        /// <param name="AddSpace"></param>
        public void GP_MatrixSetMatComboList(SAPbouiCOM.Column fCombo, string fSQL, string AndLine, string AddSpace)
        {
            SAPbobsCOM.Recordset fRecordset = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                fRecordset.DoQuery(fSQL);

                if (AddSpace == "Y")
                {
                    fCombo.ValidValues.Add("", "");
                }
                while (!fRecordset.EoF)
                {
                    fCombo.ValidValues.Add(fRecordset.Fields.Item(0).Value, fRecordset.Fields.Item(1).Value);
                    fRecordset.MoveNext();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(fRecordset);
            }
        }

        /// <summary>
        /// 콤보박스 삭제
        /// </summary>
        /// <param name="fCombo"></param>
        public void GP_MatrixRemoveMatComboList(SAPbouiCOM.Column fCombo)
        {
            int i;
            try
            {
                i = fCombo.ValidValues.Count;
                while (fCombo.ValidValues.Count > 0)
                {
                    fCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    i--;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 컨트롤 설정
        /// </summary>
        /// <param name="oForm">화면 Form</param>
        /// <param name="sItem">컨트롤명</param>
        public void AutoManaged(SAPbouiCOM.Form oForm, string sItem)
        {
            int loopCount;
            string[] ItemString = sItem.Split(',');

            oForm.AutoManaged = true;

            try
            {
                for (loopCount = 0; loopCount < ItemString.Length; loopCount++)
                {
                    oForm.Items.Item(ItemString[loopCount]).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //2:Add
                    oForm.Items.Item(ItemString[loopCount]).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //4:Find
                    oForm.Items.Item(ItemString[loopCount]).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //1:Ok
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 테이블의 내용중 현재 입력값이 존재하는지를 체크
        /// 주의 : 만약 컬럼이 숫자타입일경우가 아니면 Key_Str의 앞뒤에 "'"을 붙여 주어야 한다
        /// </summary>
        /// <param name="Tablename">테이블이름</param>
        /// <param name="ColumnName">컬럼이름</param>
        /// <param name="Key_Str">존재를 확인해야 하는키값</param>
        /// <param name="And_Line">컬럼의 데이터 타입</param>
        /// <returns></returns>
        public bool Value_ChkYn(string Tablename, string ColumnName, string Key_Str, string And_Line)
        {
            bool functionReturnValue = false;
            string sSQL;
            int Count_Chk;

            SAPbobsCOM.Recordset s_Recordset = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                if (!string.IsNullOrEmpty(Key_Str))
                {
                    sSQL = "SELECT count(*) FROM " + Tablename + " Where " + ColumnName + "=" + Convert.ToString(Key_Str);
                    if (!string.IsNullOrEmpty(And_Line))
                    {
                        sSQL += And_Line;
                    }
                    s_Recordset.DoQuery(sSQL);

                    //데이터의 존재유무 확인
                    Count_Chk = s_Recordset.Fields.Item(0).Value;

                    if (Count_Chk > 0)
                    {
                        //기존에 같은 키값으로 데이터 존재
                        functionReturnValue = false;
                    }
                    else
                    {
                        //존재하지 않는값
                        functionReturnValue = true;
                    }
                }
                else
                {
                    functionReturnValue = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(s_Recordset);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 사원정보 조회
        /// </summary>
        /// <param name="EmpCode">사원번호</param>
        /// <returns></returns>
        public ZPAY_g_EmpID Get_EmpID_InFo(string EmpCode)
        {
            //ZPAY_g_EmpID functionReturnValue = default(ZPAY_g_EmpID);
            ///// 사원순번 조회  /
            //ZPAY_g_EmpID F_EmpID = default(ZPAY_g_EmpID);

            ZPAY_g_EmpID F_EmpID = new ZPAY_g_EmpID();

            //SAPbobsCOM.Recordset Rs = new SAPbobsCOM.Recordset();
            string Sql;

            SAPbobsCOM.Recordset Rs = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Sql = "  SELECT    T0.U_EmpId,"; //사원순번
                Sql += "           T0.U_FullName,"; //사원명
                Sql += "           T0.Code,"; //사원번호
                Sql += "           T0.U_CLTCOD,"; //사업장
                Sql += "           T0.U_TeamCode,"; //부서
                Sql += "           T0.U_RspCode,"; //담당
                Sql += "           T0.U_ClsCode,"; //반
                Sql += "           Substring(replace(Convert(VarChar(10), T0.U_StartDat, 20), '-', ''), 1, 8) AS INPDAT,"; //입사일자
                Sql += "           Substring(replace(Convert(VarChar(10), T0.U_TermDate , 20), '-', ''), 1, 8) AS OUTDAT,"; //퇴사일자
                Sql += "           Substring(replace(Convert(VarChar(10), T0.U_GRPDAT , 20), '-', ''), 1, 8) AS GRPDAT,"; //그룹입사일
                //Sql = Sql & " Substring(replace(Convert(VarChar(10), T0.U_BALYMD,  20), '-', ''), 1, 8) AS BALYMD,"            '//최종발령일자
                //Sql = Sql & " T0.U_BALCOD,"                            '//최종발령부서
                Sql += "           T0.U_JIGTYP,"; //직원구분
                Sql += "           T2.posID,"; //직위(직책)코드
                Sql += "           T0.U_HOBONG ,"; //호봉
                Sql += "           T0.U_STDAMT ,"; //급여기본금
                Sql += "           T0.U_PAYTYP,"; //급여형태
                Sql += "           T0.U_PAYSEL ,"; //급여지급대상
                Sql += "           T0.U_GBHSEL ,"; //고용보험여부
                Sql += "           T0.U_govid ,"; //주민번호
                Sql += "           T0.U_sex ,"; //성별
                Sql += "           Substring(replace(Convert(VarChar(10), T0.U_RETDAT,  20), '-', ''), 1, 8) AS RETDAT,"; //중간정산일
                Sql += "           T0.U_JIGCOD,"; //직급코드
                Sql += "           (Case T0.U_BAEWOO When 'Y' then 1 else 0 end) AS U_BAEWOO,"; //배우자
                Sql += "           ISNULL(T0.U_BUYNSU, 0) AS U_BUYNSU,"; //부양가족
                Sql += "           ISNULL(T0.U_DAGYSU, 0) AS U_DAGYSU,"; //다자녀
                Sql += "           ISNULL((Select Convert(Char(8),MAX(Dateadd(dd, 1, U_ENDRET)),112) From [@PH_PY115A] Where U_MSTCOD = T0.Code), Convert(Char(8),Isnull(U_RetDat,U_STARTDAT),112)) As ENDRET ";
                Sql += " FROM      [@PH_PY001A] T0";
                Sql += "           LEFT JOIN";
                Sql += "           [OUDP] T1";
                Sql += "               ON T0.U_TeamCode = T1.Code";
                Sql += "           LEFT JOIN";
                Sql += "           [OHPS] T2";
                Sql += "               ON T0.U_Position = T2.PosID";
                //    Sql = Sql & " LEFT JOIN   (SELECT T0.*, T1.U_RelCd"F_EmpID
                //    Sql = Sql & " FROM [@PH_PY001A] T0 INNER JOIN [@PS_HR200L] T1 ON T0.U_PAYTYP = T1.U_Code AND T1.Code = 'P132') T3 ON T0.U_MSTCOD = T3.Code"
                Sql += " WHERE     T0.Code = '" + EmpCode + "'";
                Sql += " ORDER BY  T0.Code";

                Rs.DoQuery(Sql);

                if (Rs.RecordCount == 0)
                {
                    F_EmpID.EmpID = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.MSTNAM = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.MSTCOD = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.CLTCOD = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.TeamCode = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.RspCode = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.CLTCOD = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.StartDate = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.TermDate = new String(' ', 0); //Strings.Space(0);
                    //F_EmpID.BALYMD = Strings.Space(0);
                    //F_EmpID.BALCOD = Strings.Space(0);
                    F_EmpID.JIGTYP = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.PAYTYP = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.PAYSEL = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.Position = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.HOBONG = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.STDAMT = 0;
                    F_EmpID.GBHSEL = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.PERNBR = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.Sex = "";
                    F_EmpID.RETDAT = "";
                    F_EmpID.JIGCOD = "";
                    F_EmpID.GONCNT = 0;
                    F_EmpID.DAGYSU = 0;
                    F_EmpID.GRPDAT = new String(' ', 0); //Strings.Space(0);
                    F_EmpID.ENDRET = new String(' ', 0); //Strings.Space(0);
                }
                else
                {
                    while (!Rs.EoF)
                    {
                        F_EmpID.EmpID = Rs.Fields.Item("U_EmpID").Value; //사원순번
                        F_EmpID.MSTNAM = Rs.Fields.Item("U_FullName").Value; //사원명
                        F_EmpID.MSTCOD = Rs.Fields.Item("Code").Value; //사원코드
                        F_EmpID.CLTCOD = Rs.Fields.Item("U_CLTCOD").Value; //사업장
                        F_EmpID.TeamCode = Rs.Fields.Item("U_TeamCode").Value; //부서
                        F_EmpID.RspCode = Rs.Fields.Item("U_RspCode").Value; //담당
                        F_EmpID.ClsCode = Rs.Fields.Item("U_ClsCode").Value; //반
                        F_EmpID.StartDate = Rs.Fields.Item("INPDAT").Value; //입사일자
                        F_EmpID.TermDate = Rs.Fields.Item("OUTDAT").Value; //퇴사일자
                        //.BALYMD = Rs.Fields("U_BALYMD").Value       '//최종발령일자
                        //.BALCOD = Rs.Fields("U_BALCOD").Value       '//최종발령부서
                        F_EmpID.JIGTYP = Rs.Fields.Item("U_JIGTYP").Value; //직종
                        F_EmpID.Position = Rs.Fields.Item("PosID").Value.ToString().Trim(); //직위
                        F_EmpID.HOBONG = Rs.Fields.Item("U_Hobong").Value; //호봉
                        F_EmpID.STDAMT = Rs.Fields.Item("U_STDAMT").Value; //기본급
                        F_EmpID.PAYTYP = Rs.Fields.Item("U_PAYTYP").Value; //급여형태
                        F_EmpID.PAYSEL = Rs.Fields.Item("U_PAYSEL").Value; //급여지급일구분
                        F_EmpID.GBHSEL = Rs.Fields.Item("U_GBHSEL").Value; //고용보험납입여부
                        F_EmpID.PERNBR = Rs.Fields.Item("U_govid").Value; //주민번호
                        F_EmpID.Sex = Rs.Fields.Item("U_SEX").Value; //성별
                        F_EmpID.RETDAT = Rs.Fields.Item("RETDAT").Value; //중도정산일자
                        F_EmpID.JIGCOD = Rs.Fields.Item("U_JIGCOD").Value; //직급
                        F_EmpID.GONCNT = Convert.ToInt16(1 + Rs.Fields.Item("U_BAEWOO").Value + Rs.Fields.Item("U_BUYNSU").Value); //부양가족
                        F_EmpID.DAGYSU = Convert.ToInt16(Rs.Fields.Item("U_DAGYSU").Value); //다자녀공제
                        F_EmpID.GRPDAT = Rs.Fields.Item("GRPDAT").Value; //그룹입사일자
                        F_EmpID.ENDRET = Rs.Fields.Item("ENDRET").Value; //퇴충기산일

                        Rs.MoveNext();
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Rs);
            }

            return F_EmpID;
        }

        /// <summary>
        /// 작업연월 잠김여부 체크
        /// </summary>
        /// <param name="sJOBYMM"></param>
        /// <param name="sJOBTYP"></param>
        /// <param name="sJOBGBN"></param>
        /// <param name="sPAYSEL"></param>
        /// <returns></returns>
        public bool Get_PayLockInfo(string sJOBYMM, string sJOBTYP, string sJOBGBN, string sPAYSEL)
        {
            bool functionReturnValue = false;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  ISNULL(U_ENDCHK, 'N') ";
                sQry += " FROM    [@ZPY307L] ";
                sQry += " WHERE   Code = '" + sJOBYMM.Substring(0, 4) + "' ";
                sQry += "         AND    U_JOBYMM = '" + sJOBYMM + "' ";
                if (sJOBTYP.Trim() != "%" && !string.IsNullOrEmpty(sJOBTYP.Trim()))
                {
                    sQry += "     AND   (CASE WHEN U_JOBTYP = '%' THEN '" + sJOBTYP + "' ELSE U_JOBTYP END) LIKE '" + sJOBTYP + "' ";
                }
                if (sJOBGBN.Trim() != "%" && !string.IsNullOrEmpty(sJOBGBN.Trim()))
                {
                    sQry += "     AND   (CASE WHEN U_JOBGBN = '%' THEN '" + sJOBGBN + "' ELSE U_JOBTYP END) LIKE '" + sJOBGBN + "' ";
                }
                if (sPAYSEL.Trim() != "%" && !string.IsNullOrEmpty(sPAYSEL.Trim()))
                {
                    sQry += "     AND   (CASE WHEN U_PAYSEL = '%' THEN '" + sPAYSEL + "' ELSE U_JOBTYP END) LIKE '" + sPAYSEL + "' ";
                }

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    functionReturnValue = false;
                }
                else if (oRecordSet.Fields.Item(0).Value == "N")
                {
                    functionReturnValue = false;
                }
                else
                {
                    functionReturnValue = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            
            return functionReturnValue;
        }

        /// <summary>
        /// 사업자 번호 체크
        /// (실행 테스트 필요, 테스트 완료 후 해당 주석 라인 삭제)
        /// </summary>
        /// <param name="strNo"></param>
        /// <returns></returns>
        public bool TaxNoCheck(string strNo)
        {
            const byte COMPNO_LEN = 10; //사업자번호의 길이
            bool blnRet; //결과값
            byte[] aryNo = new byte[COMPNO_LEN + 1]; //문자열 배열
            int bytCntNo; //루프변수
            short intMod; //나머지숫자
            short intInt; //소수점이하 절사값
            short intSub; //계산결과 
            string BUSNBR; //사업자번호

            BUSNBR = strNo.Replace("-", "");
            
            if (BUSNBR.Trim().Length == COMPNO_LEN) //사업자번호의 길이가 10자리라면
            {
                //루프를 돌면서 바이트배열을 만든다
                for (bytCntNo = 1; bytCntNo <= COMPNO_LEN; bytCntNo++)
                {
                    aryNo[bytCntNo] = Convert.ToByte(BUSNBR.Substring(bytCntNo - 1, 1));
                }

                //나머지 숫자를 구한다
                intMod = Convert.ToInt16(((aryNo[1] * 1) + (aryNo[2] * 3) + (aryNo[3] * 7) + (aryNo[4] * 1) + (aryNo[5] * 3) + (aryNo[6] * 7) + (aryNo[7] * 1) + (aryNo[8] * 3)) % COMPNO_LEN);
                //소숫점이하를 절사하여 구한다
                intInt = Convert.ToInt16(aryNo[9] * 5 / COMPNO_LEN);
                //계산결과를 구한다
                intSub = Convert.ToInt16((aryNo[9] * 5) - (intInt * 10));

                intSub = Convert.ToInt16((intMod + intInt + intSub) % 10);

                intSub = Convert.ToInt16((intSub == 0) ? 10 : intSub);

                //체크섬을 확인하여 진위를 판별한다
                blnRet = (aryNo[COMPNO_LEN] == (COMPNO_LEN - intSub));
            }
            else
            {
                blnRet = false;
            }
            //결과를 대입한다
            return blnRet;
        }

        /// <summary>
        /// 급여나 기타 금액 계산시 끝단위 처리
        /// (실행 테스트 필요, 테스트 완료 후 해당 주석 라인 삭제)
        /// </summary>
        /// <param name="Dub">금액</param>
        /// <param name="oPnt">비율</param>
        /// <param name="Rtype">끝전 처리 방법(R:반올림, F:절사, C:올림)</param>
        /// <returns></returns>
        public int RInt(double Dub, short oPnt, string Rtype)
        {
            int functionReturnValue = 0;
            double Rub = 0;
            double Cub = 0;
            short Pnt;

            if (Dub == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            Pnt = Convert.ToInt16(oPnt);

            switch (Pnt)
            {
                case 1: // 1원
                    Rub = 0.5;
                    Cub = 0.9;
                    break;
                case 10: // 10원
                    Rub = 5;
                    Cub = 9;
                    break;
                case 100:
                    Rub = 50;
                    Cub = 90;
                    break;
                case 1000:
                    Rub = 500;
                    Cub = 900;
                    break;
            }

            switch (Rtype.Trim())
            {//기존VB6.0 소스코드의 Int 함수는 정수부분만 return 하기 때문에 결과적으로는 모든 수를 버림하는 결과가 나옴(2019.12.09 송명규)
                case "R":
                    functionReturnValue = Convert.ToInt32(Math.Truncate((Dub + Rub) / Pnt) * Pnt);
                    //Int((Dub + Rub) / Pnt) * Pnt
                    break;
                case "C":
                    functionReturnValue = Convert.ToInt32(Math.Truncate((Dub + Cub) / Pnt) * Pnt);
                    //Int((Dub + Cub) / Pnt) * Pnt
                    break;
                case "F":
                    functionReturnValue = Convert.ToInt32(Math.Truncate(Dub / Pnt) * Pnt);
                    //Int(Dub / Pnt) * Pnt
                    break;
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 근속기간 등을 계산하는 함수(시작일자를 기준으로 날짜를 1년, 1개월, 1일씩 더해서 종료일자가 될때까지 카운트해서 계산함)
        /// (실행 테스트 필요, 테스트 완료 후 해당 주석 라인 삭제)
        /// </summary>
        /// <param name="STRDAT"></param>
        /// <param name="ENDDAT"></param>
        public void Term2(string STRDAT, string ENDDAT)
        {
            DateTime CHKDAY = new DateTime();
            short TempCnt;

            PSH_Globals.ZPAY_GBL_GNSYER = 0;
            PSH_Globals.ZPAY_GBL_GNMYER = 0;
            PSH_Globals.ZPAY_GBL_GNSMON = 0;
            PSH_Globals.ZPAY_GBL_GNMMON = 0;
            PSH_Globals.ZPAY_GBL_GNSDAY = 0;
            PSH_Globals.ZPAY_GBL_GNMDAY = 0;
            
            DateTime ENDDAT1 = DateTime.ParseExact(ENDDAT, "yyyyMMdd", null).AddDays(1); //1일 추가
            DateTime CHKDAY1 = DateTime.ParseExact(STRDAT, "yyyyMMdd", null);

            //근속년수 체크
            TempCnt = 0;

            while (!(CHKDAY > ENDDAT1))
            {
                TempCnt = (short)(TempCnt + 1);
                CHKDAY = CHKDAY1.AddYears(TempCnt);
            }
            PSH_Globals.ZPAY_GBL_GNSYER = (short)(TempCnt - 1);
            CHKDAY1 = CHKDAY1.AddYears(PSH_Globals.ZPAY_GBL_GNSYER);
            CHKDAY = CHKDAY1;

            // 근속월수 체크
            TempCnt = 0;
            while (!(CHKDAY > ENDDAT1))
            {
                TempCnt = (short)(TempCnt + 1);
                CHKDAY = CHKDAY1.AddMonths(TempCnt);
            }
            PSH_Globals.ZPAY_GBL_GNSMON = (short)(TempCnt - 1);
            CHKDAY1 = CHKDAY1.AddMonths(PSH_Globals.ZPAY_GBL_GNSMON);
            CHKDAY = CHKDAY1;

            // 근속일수 체크
            TempCnt = 0;
            while (!(CHKDAY > ENDDAT1))
            {
                TempCnt = (short)(TempCnt + 1);
                CHKDAY = CHKDAY1.AddDays(TempCnt);
            }
            PSH_Globals.ZPAY_GBL_GNSDAY = (short)(TempCnt - 1);
            CHKDAY = CHKDAY1.AddDays(PSH_Globals.ZPAY_GBL_GNSDAY);
            
            PSH_Globals.ZPAY_GBL_GNMYER = PSH_Globals.ZPAY_GBL_GNSYER; //근속연수
            PSH_Globals.ZPAY_GBL_GNMMON = (short)(PSH_Globals.ZPAY_GBL_GNSYER * 12 + PSH_Globals.ZPAY_GBL_GNSMON); //근속월수
        }

        /// <summary>
        /// 소수점 절사
        /// </summary>
        /// <param name="Dub">절사대상</param>
        /// <param name="Pnt">비율</param>
        /// <returns></returns>
        public double IInt(double Dub, double Pnt)
        {
            string SDub;
            string[] arrSDub;
            double TDub;
            double Tub;

            Tub = (Dub >= 0 ? (Dub / Pnt) : (Dub / Pnt * -1)); //13자리 이상의 숫자일 경우 Pnt를 2를 줘서 숫자를 반으로 줄임(VB6.0에서 13자리 이상의 수를 소수점 절사하기위한 알고리즘으로 판단됨)-SongMG
            SDub = Tub.ToString("0000000000000.000000");

            arrSDub = SDub.Split('.');

            TDub = Convert.ToDouble(arrSDub[0].ToString());

            return (Dub >= 0 ? (TDub * Pnt) : (TDub * Pnt * -1)); //반으로 줄인 수를 다시 Pnt인 2를 곱해서 원상복귀 시켜서 리턴
        }

        /// <summary>
        /// 갑근세와 주민세 계산
        /// (실행 테스트 필요, 테스트 완료 후 해당 주석 라인 삭제)
        /// </summary>
        /// <param name="GABGUN">리턴받을 갑근세 ref 변수</param>
        /// <param name="JUMINN">리턴받을 주민세 ref 변수</param>
        /// <param name="oINCOME"></param>
        /// <param name="oInWON"></param>
        /// <param name="oChlWON"></param>
        /// <param name="JOBYMM"></param>
        /// <param name="oKUKAMT"></param>
        /// <param name="PAY_001"></param>
        /// <returns></returns>
        public object Get_GabGunSe_Table(ref double GABGUN, ref double JUMINN, double oINCOME, short oInWON, short oChlWON, string JOBYMM, double oKUKAMT, string PAY_001)
        {
            object functionReturnValue = null;
            string sQry;

            double WK_INCOME;
            double WK_GULTAX = 0;

            SAPbobsCOM.Recordset Rs = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            GABGUN = 0;
            JUMINN = 0;

            try
            {
                //총지급액
                if (oINCOME <= 0)
                {
                    functionReturnValue = "과세금액이 0보다 작거나 같습니다. 확인하여 주세요.";
                    return functionReturnValue;
                }

                WK_INCOME = oINCOME;

                if (Convert.ToInt32(JOBYMM) <= 201201)
                {
                    // 1000만원초과시
                    if (oINCOME > 10000000)
                    {
                        GABGUN = IInt(((oINCOME - 10000000) * 0.95) * 0.35, 1);
                        WK_INCOME = 10000000;
                    }
                }
                else
                {
                    if (oINCOME > 28000000)
                    {
                        // 2800만원초과시
                        GABGUN = 5985000 + IInt((oINCOME - 28000000) * 0.95 * 0.38, 1);
                        WK_INCOME = 10000000;
                    }
                    else if (oINCOME > 10000000)
                    {
                        // 1000만원초과시
                        GABGUN = IInt((oINCOME - 10000000) * 0.95 * 0.35, 1);
                        WK_INCOME = 10000000;
                    }
                }

                if (Convert.ToInt32(JOBYMM) >= 201101 && oChlWON > 0)
                {
                    oInWON = (short)(oInWON + oChlWON - 1);
                    oChlWON = 0;
                }

                // 간이세액조견표 등록된 테이블값 참조
                sQry = " SELECT TOP 1 ISNULL(T0.U_CODAVR,0) AS U_CODAVR,";
                sQry += "       ISNULL(CASE WHEN " + oInWON + " <= 1  THEN U_BY01ST";
                sQry += "                   WHEN " + oInWON + "  = 2  THEN U_BY02ST";
                sQry += "                   WHEN " + oInWON + "  = 3  AND " + oChlWON + "  < 2 THEN U_BY03ST";
                sQry += "                   WHEN " + oInWON + "  = 3  AND " + oChlWON + " >= 2 THEN U_BY03DJ";
                sQry += "                   WHEN " + oInWON + "  = 4  AND " + oChlWON + "  < 2 THEN U_BY04ST";
                sQry += "                   WHEN " + oInWON + "  = 4  AND " + oChlWON + " >= 2 THEN U_BY04DJ";
                sQry += "                   WHEN " + oInWON + "  = 5  AND " + oChlWON + "  < 2 THEN U_BY05ST";
                sQry += "                   WHEN " + oInWON + "  = 5  AND " + oChlWON + " >= 2 THEN U_BY05DJ";
                sQry += "                   WHEN " + oInWON + "  = 6  AND " + oChlWON + "  < 2 THEN U_BY06ST";
                sQry += "                   WHEN " + oInWON + "  = 6  AND " + oChlWON + " >= 2 THEN U_BY06DJ";
                sQry += "                   WHEN " + oInWON + "  = 7  AND " + oChlWON + "  < 2 THEN U_BY07ST";
                sQry += "                   WHEN " + oInWON + "  = 7  AND " + oChlWON + " >= 2 THEN U_BY07DJ";
                sQry += "                   WHEN " + oInWON + "  = 8  AND " + oChlWON + "  < 2 THEN U_BY08ST";
                sQry += "                   WHEN " + oInWON + "  = 8  AND " + oChlWON + " >= 2 THEN U_BY08DJ";
                sQry += "                   WHEN " + oInWON + "  = 9  AND " + oChlWON + "  < 2 THEN U_BY09ST";
                sQry += "                   WHEN " + oInWON + "  = 9  AND " + oChlWON + " >= 2 THEN U_BY09DJ";
                sQry += "                   WHEN " + oInWON + "  = 10 AND " + oChlWON + "  < 2 THEN U_BY10ST";
                sQry += "                   WHEN " + oInWON + "  = 10 AND " + oChlWON + " >= 2 THEN U_BY10DJ";
                sQry += "                   WHEN " + oInWON + " >= 11 AND " + oChlWON + "  < 2 THEN U_BY11ST";
                sQry += "                   WHEN " + oInWON + " >= 11 AND " + oChlWON + " >= 2 THEN U_BY11DJ";
                sQry += "                   ELSE 0 END, 0) AS U_GABGUB ";
                sQry += " FROM [@ZPY301L] T0 WHERE   T0.CODE <= '" + JOBYMM + "'";
                sQry += " AND     T0.U_CODFRS <= " + WK_INCOME + " AND     T0.U_CODTOM >  " + WK_INCOME + "";
                sQry += " ORDER BY T0.Code Desc";

                Rs.DoQuery(sQry);

                if (Rs.RecordCount != 0)
                {
                    WK_GULTAX = Rs.Fields.Item("U_GABGUB").Value;
                }

                //갑근세
                GABGUN = IInt(GABGUN + WK_GULTAX, 1);

                if (GABGUN < 1000)
                {
                    GABGUN = 0;
                }

                //지방소득세(주민세)
                JUMINN = IInt(GABGUN * 0.1, 1);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Rs);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 갑근세와 주민세 계산(공제인원수가 4명이 아닐경우 호출)
        /// (실행 테스트 필요, 테스트 완료 후 해당 주석 라인 삭제)
        /// </summary>
        /// <param name="GABGUN"></param>
        /// <param name="JUMINN"></param>
        /// <param name="oINCOME"></param>
        /// <param name="oInWON"></param>
        /// <param name="oChlWON"></param>
        /// <param name="JOBYMM"></param>
        /// <param name="oKUKAMT"></param>
        /// <param name="PAY_001"></param>
        /// <returns></returns>
        public object Get_GabGunSe(ref double GABGUN, ref double JUMINN, double oINCOME, short oInWON, short oChlWON, string JOBYMM, double oKUKAMT, string PAY_001)
        {
            object functionReturnValue = null;
            string sQry;

            double WS_INCOME;
            double WK_INCOME;
            double WK_GNLOSD;
            double WK_SANTAX;
            double WK_TAXGON;
            double WK_KUKAMT;
            double WK_GULTAX;

            SAPbobsCOM.Recordset Rs = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            GABGUN = 0;
            JUMINN = 0;

            try
            {
                // 총지급액
                if (oINCOME <= 0)
                {
                    functionReturnValue = "과세금액이 0보다 작거나 같습니다. 확인하여 주세요.";
                    return functionReturnValue;
                }

                //  간이세액조견표 구간 평균값을 사용할 경우
                if (PAY_001 == "2" || PAY_001 == "3")
                {
                    sQry = "  SELECT TOP 1 ISNULL(T0.U_CODAVR,0) AS U_CODAVR FROM [@ZPY301L] T0 WHERE   T0.CODE <= '" + JOBYMM + "'";
                    sQry += " AND     T0.U_CODFRS <= " + oINCOME + " AND     T0.U_CODTOM >  " + oINCOME + "";
                    sQry += " ORDER BY T0.Code Desc";
                    Rs.DoQuery(sQry);
                    if (Rs.RecordCount != 0)
                    {
                        oINCOME = Rs.Fields.Item("U_CODAVR").Value;
                        oKUKAMT = oINCOME;
                        WS_INCOME = oINCOME;
                    }
                }

                WK_INCOME = oINCOME;
                WS_INCOME = oINCOME;

                if (Convert.ToInt32(JOBYMM) <= 201201)
                {
                    // 1000만원초과시
                    if (oINCOME > 10000000)
                    {
                        GABGUN = IInt((oINCOME - 10000000) * 0.95 * 0.35, 1);
                        WK_INCOME = 10000000;
                        WS_INCOME = 10000000;
                    }
                }
                else
                {
                    if (oINCOME > 28000000)
                    {
                        // 2800만원초과시
                        GABGUN = 5985000 + IInt((oINCOME - 28000000) * 0.95 * 0.38, 1);
                    }
                    else if (oINCOME > 10000000)
                    {
                        // 1000만원초과시
                        GABGUN = IInt((oINCOME - 10000000) * 0.95 * 0.35, 1);
                    }
                }

                // 2008년까지만(기존업체 관리직 변경어려움)
                //   If Left(JOBYMM, 4) = "2008" Then
                //        Select Case Trim(MDC_COMpanyGubun)
                //        Case "OBS"
                //        WS_INCOME = oINCOME
                //        End Select
                //   End If

                WK_INCOME *= 12;
                // 근로소득공제(2007.01시행)
                if (Convert.ToInt16(JOBYMM.Substring(0, 4)) <= 2008)
                {
                    //(근로소득: 500만원이하              전액공제
                    //           500만원초과~1500만원이하  500만원+(근로소득- 500만원)*50%
                    //          1500만원초과~3000만원이하 1000만원+(근로소득-1500만원)*15%)
                    //          4500만원이하              1225만원+(근로소득-3000만원)*10%)
                    //          4500만원초과              1375만원+(근로소득-4500만원)* 5%) 한도없슴
                    if (WK_INCOME <= 5000000)
                    {
                        WK_GNLOSD = WK_INCOME;
                    }
                    else if (WK_INCOME <= 15000000)
                    {
                        WK_GNLOSD = 5000000 + (WK_INCOME - 5000000) * 0.5;
                        //3000
                    }
                    else if (WK_INCOME <= 30000000)
                    {
                        WK_GNLOSD = 10000000 + (WK_INCOME - 15000000) * 0.15;
                        //4500
                    }
                    else if (WK_INCOME <= 45000000)
                    {
                        WK_GNLOSD = 12250000 + (WK_INCOME - 30000000) * 0.1;
                    }
                    else
                    {
                        WK_GNLOSD = 13750000 + (WK_INCOME - 45000000) * 0.05;
                    }
                }
                else
                {
                    //2009년 근로소득공제금액 개정
                    //(근로소득: 500만원이하              전액*80%
                    //           500만원초과~1500만원이하  400만원+(근로소득- 500만원)*50%
                    //          1500만원초과~3000만원이하  900만원+(근로소득-1500만원)*15%)
                    //          4500만원이하              1125만원+(근로소득-3000만원)*10%)
                    //          4500만원초과              1275만원+(근로소득-4500만원)* 5%) 한도없슴
                    if (WK_INCOME <= 5000000)
                    {
                        WK_GNLOSD = WK_INCOME;
                    }
                    else if (WK_INCOME <= 15000000)
                    {
                        WK_GNLOSD = 4000000 + (WK_INCOME - 5000000) * 0.5;
                        //3000
                    }
                    else if (WK_INCOME <= 30000000)
                    {
                        WK_GNLOSD = 9000000 + (WK_INCOME - 15000000) * 0.15;
                        //4500
                    }
                    else if (WK_INCOME <= 45000000)
                    {
                        WK_GNLOSD = 11250000 + (WK_INCOME - 30000000) * 0.1;
                    }
                    else
                    {
                        WK_GNLOSD = 12750000 + (WK_INCOME - 45000000) * 0.05;
                    }
                }

                // 근로소득금액 ( 근로소득-근로소득공제 ) /
                WK_INCOME -= WK_GNLOSD;
                // 기본공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                if (Convert.ToInt32(JOBYMM) <= 200812)
                {
                    //  인적공제 1인당 100만원
                    //  WK_INCOME = WK_INCOME - 1000000                   '/ 1.본        인 /
                    WK_INCOME -= (1000000 * oInWON);
                    // 2.부양가족공제 /
                }
                else
                {
                    //  인적공제 1인당 150만원
                    //  WK_INCOME = WK_INCOME - 1500000                   '/ 1.본        인 /
                    WK_INCOME -= (1500000 * oInWON);
                    // 2.부양가족공제 /
                }

                //(2007.01시행 변경내용 //////////////////////////////////////////////////////////////////////
                // 소수공제자추가공제 폐지
                // 다자녀추가공제 신설: 20세이하자녀가 2인 50만원, 2인초과 50만원 +(2인초과인원수*100만원)
                //////////////////////////////////////////////////////////////////////////////////////////////
                // 소수인원추가공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                // 소수공제 1인 100만원, 2인 50만원
                //   Select Case (oInWON)
                //     Case 1: WK_INCOME = WK_INCOME - 1000000
                //     Case 2: WK_INCOME = WK_INCOME - 500000
                //   End Select
                // 다자녀추가공제 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                if (oChlWON > 1 && oInWON > 2)
                {
                    if (oChlWON <= 2)
                    {
                        WK_INCOME -= 500000;
                    }
                    else
                    {
                        WK_INCOME -= 500000;
                        // 2009.05월분 이전에 다자녀 공제 2인이상 추가인원수 공제했던거는 그대로
                        if (PAY_001 == "1" || PAY_001 == "2")
                        {
                            WK_INCOME -= (1000000 * (oChlWON - 2));
                        }
                    }
                }

                // 특별공제(2인이하인경우1,200,000 3인이상인경우 2,400,000)
                // 특별공제-2008년4월부터변경됨 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                // (2인이하인경우: 1,200,000 => 100만원과 연간급여액의 25/1000해당하는 금액의 합계액
                // (3인이상인경우: 2,400,000 => 240만원과 연간급여액의 5/100해당하는 금액의 합계액+ 연간급여액에서 4천만원초과금액의 5/100
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                if (Convert.ToInt32(JOBYMM) <= 200712)
                {
                    if (oInWON <= 2)
                    {
                        WK_INCOME -= (1000000 + (WS_INCOME * 12 * 2.5 / 100));
                    }
                    else
                    {
                        WK_INCOME -= (2400000 + (WS_INCOME * 12 * 5 / 100));
                    }
                }
                else
                {
                    if (oInWON <= 2)
                    {
                        WK_INCOME -= (1100000 + (WS_INCOME * 12 * 2.5 / 100));
                    }
                    else
                    {
                        WK_INCOME -= (2500000 + (WS_INCOME * 12 * 5 / 100));
                        if ((WS_INCOME * 12) > 40000000)
                        {
                            WK_INCOME -= ((WS_INCOME * 12 - 40000000) * 5 / 100);
                        }
                    }
                }

                // 연금보험료공제(2008.03월까지는 국민연금등급제, 2008년04월부터 국민연금보수월액제
                if (Convert.ToInt32(JOBYMM) <= 200712)
                {
                    // (국민연금조견표의 각출료 *12)
                    sQry = "  SELECT  T0.U_EMPAMT, T0.U_FROM, T0.U_TO";
                    sQry += " FROM [@ZPY102L] T0 INNER JOIN [@ZPY102H] T1 ON T0.Code = T1.Code";
                    sQry += " WHERE T1.Code <= '" + JOBYMM + "'";
                    sQry += " AND  T0.U_FROM <= " + WS_INCOME + "";
                    sQry += " AND T0.U_TO > " + WS_INCOME + "";
                    sQry += " ORDER BY T1.Code Desc";
                    Rs.DoQuery(sQry);
                    if (Rs.RecordCount != 0)
                    {
                        WK_INCOME = IInt(WK_INCOME - (Rs.Fields.Item("U_EMPAMT").Value * 12), 1);
                    }
                    // 2008년 4월부터
                }
                else
                {
                    sQry = "  SELECT TOP 1 U_EMPRAT, U_FROM, U_TO FROM [@ZPY102H] ";
                    sQry += " WHERE CODE >= '200804' ORDER BY CODE DESC";
                    Rs.DoQuery(sQry);
                    if (Rs.RecordCount != 0)
                    {
                        if (oKUKAMT < Rs.Fields.Item("U_FROM").Value)
                        {
                            WK_KUKAMT = Rs.Fields.Item("U_FROM").Value;
                        }
                        else if (Rs.Fields.Item("U_TO").Value > 0 && oKUKAMT > Rs.Fields.Item("U_TO").Value)
                        {
                            WK_KUKAMT = Rs.Fields.Item("U_TO").Value;
                        }
                        else
                        {
                            WK_KUKAMT = oKUKAMT;
                        }

                        WK_KUKAMT = IInt(WK_KUKAMT * 12 * Rs.Fields.Item("U_EMPRAT").Value / 100, 1);

                        WK_INCOME -= WK_KUKAMT;
                    }
                }
                // 과세표준 ( 근로소득금액 - 인적공제 - 특별공제 - 기타소득공제 ) /
                if (WK_INCOME < 0)
                {
                    WK_INCOME = 0;
                }
                // 산출세액
                if (Convert.ToInt32(JOBYMM) <= 200812)
                {
                    //2008년도
                    //(과세표준:1200만원이하               과세표준*8%
                    //          1200만원초과~4600만원이하  과세표준*17%-  96만원
                    //          4600만원초과~8800만원이하  과세표준*26%- 674만원
                    //          8800만원초과               과세표준*35%-1766만원)
                    if (WK_INCOME <= 12000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.08 - 0;
                    }
                    else if (WK_INCOME <= 46000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.17 - 1080000;
                    }
                    else if (WK_INCOME <= 88000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.26 - 5220000;
                    }
                    else
                    {
                        WK_SANTAX = WK_INCOME * 0.35 - 13140000;
                    }
                }
                else if (JOBYMM == "200912")
                {
                    //2009년도
                    //(과세표준:1200만원이하               과세표준*6%
                    //          1200만원초과~4600만원이하  과세표준*16%-  72만원
                    //          4600만원초과~8800만원이하  과세표준*26%- 616만원
                    //          8800만원초과               과세표준*35%-1666만원)
                    if (WK_INCOME <= 12000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.06 - 0;
                    }
                    else if (WK_INCOME <= 46000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.16 - 1200000;
                    }
                    else if (WK_INCOME <= 88000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.25 - 5340000;
                    }
                    else
                    {
                        WK_SANTAX = WK_INCOME * 0.35 - 14140000;
                    }
                }
                else if (Convert.ToInt32(JOBYMM) <= 201201)
                {
                    //2010년도
                    //(과세표준:1200만원이하               과세표준*6%
                    //          1200만원초과~4600만원이하  과세표준*15%-  72만원
                    //          4600만원초과~8800만원이하  과세표준*24%- 582만원
                    //          8800만원초과               과세표준*35%-1590만원)
                    if (WK_INCOME <= 12000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.06 - 0;
                    }
                    else if (WK_INCOME <= 46000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.15 - 1080000;
                    }
                    else if (WK_INCOME <= 88000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.24 - 5220000;
                    }
                    else
                    {
                        WK_SANTAX = WK_INCOME * 0.35 - 14900000;
                    }
                }
                else
                {
                    //2012년도
                    //(과세표준:1200만원이하               과세표준*6%
                    //          1200만원초과~4600만원이하  과세표준*15%-  72만원
                    //          4600만원초과~8800만원이하  과세표준*24%-  582만원
                    //          8000만원초과~3억원 이하    과세표준*35%-  1590만원)
                    //          3억원 초과                 과세표준*38%-  9010만원)
                    if (WK_INCOME <= 12000000)
                    {
                        WK_SANTAX = WK_INCOME * 0.06 - 0;
                    }
                    else if (WK_INCOME <= 46000000)
                    {
                        WK_SANTAX = (WK_INCOME - 12000000) * 0.15 + 720000;
                    }
                    else if (WK_INCOME <= 88000000)
                    {
                        WK_SANTAX = (WK_INCOME - 46000000) * 0.24 + 5820000;
                    }
                    else if (WK_INCOME <= 300000000)
                    {
                        WK_SANTAX = (WK_INCOME - 88000000) * 0.35 + 15900000;
                    }
                    else
                    {
                        WK_SANTAX = (WK_INCOME - 300000000) * 0.38 + 90100000;
                    }
                }

                WK_SANTAX = IInt(WK_SANTAX, 1);
                //  세액공제(2007.01 시행)
                //  50만원이하  산출세액 * 55%
                //  50만원초과  275000 + (산출세액-500000) * 30%
                //  세액공제한도액: 45만원한도
                if (WK_SANTAX <= 500000)
                {
                    WK_TAXGON = WK_SANTAX * 0.55;
                }
                else
                {
                    WK_TAXGON = 275000 + (WK_SANTAX - 500000) * 0.3;
                }

                WK_TAXGON = IInt(WK_TAXGON, 1);

                if (WK_TAXGON > 500000)
                {
                    WK_TAXGON = 500000;
                }
                    

                // 결정세액 ( 산출세액 - 세액공제 및 감면 ) /
                WK_GULTAX = IInt((WK_SANTAX - WK_TAXGON) / 12, 1);

                // 갑근세
                GABGUN = IInt(GABGUN + WK_GULTAX, 1);

                if (GABGUN < 1000)
                {
                    GABGUN = 0;
                }
                // 지방소득세(주민세)
                JUMINN = IInt(GABGUN * 0.1, 1);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Rs);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 근속일수 조회
        /// </summary>
        /// <param name="StrDate"></param>
        /// <param name="EndDate"></param>
        /// <returns></returns>
        public int TermDay(string StrDate, string EndDate)
        {
            DateTime STRDAT = DateTime.ParseExact(StrDate, "yyyyMMdd", null);
            DateTime ENDDAT = DateTime.ParseExact(EndDate, "yyyyMMdd", null);

            //일자 Format의 유효성확인이 필요할까? 주석처리(2019.05.10 송명규)
            //if (Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StrDate, "0000-00-00")) == false || Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(EndDate, "0000-00-00")) == false)
            //{
            //    return functionReturnValue;
            //}

            TimeSpan timeDiff = ENDDAT - STRDAT; //TimeSpan을 이용해서 일자 차이 구함
                
            return timeDiff.Days + 1; //1일을 추가하여 리턴
        }

        /// <summary>
        /// 해당월의 마지막일자 조회
        /// </summary>
        /// <param name="YMM"></param>
        /// <returns></returns>
        public string Lday(string YMM)
        {
            //기존 VB6.0 로직_S
            //object functionReturnValue = null;

            //switch (true)
            //{
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(YMM, 1, 6) + "31", "0000-00-00")):
            //        functionReturnValue = "31";
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(YMM, 1, 6) + "30", "0000-00-00")):
            //        functionReturnValue = "30";
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(YMM, 1, 6) + "29", "0000-00-00")):
            //        functionReturnValue = "29";
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(YMM, 1, 6) + "28", "0000-00-00")):
            //        functionReturnValue = "28";
            //        break;
            //    default:
            //        functionReturnValue = Strings.Space(0);
            //        break;
            //}
            //return functionReturnValue;
            //기존 VB6.0 로직_E

            return DateTime.DaysInMonth(Convert.ToInt16(YMM.Substring(0, 4)), Convert.ToInt16(YMM.Substring(4, 2))).ToString(); //한줄로 끝
        }

        /// <summary>
        /// 폴더 생성(연말정산 신고용 자료 생성시 폴더 생성, Z, R 클래스에서 사용, 신규 클래스에서는 미사용)
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public string CreateFolder(string FileName)
        {
            string functionReturnValue = null;

            Scripting.FileSystemObject fs = new Scripting.FileSystemObject();

            try
            {
                if (fs.FolderExists(FileName) == false)
                {
                    fs.CreateFolder(FileName);
                }

                functionReturnValue = "";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            
            return functionReturnValue;
        }

        #region sStr 메소드(한글을 2Byte로 변환) 미사용
        ///// <summary>
        ///// 한글을 2Byte로 변환
        ///// 사용금지
        ///// </summary>
        ///// <param name="GetStr"></param>
        ///// <returns></returns>
        //public string sStr(string GetStr)
        //{
        //    string returnValue = string.Empty;
        //    //임시주석_S
        //    //returnValue = Microsoft.VisualBasic.Strings.Left(Microsoft.VisualBasic.Strings.StrConv(GetStr, vbFromUnicode), Microsoft.VisualBasic.Strings.Len(GetStr));
        //    //returnValue = Microsoft.VisualBasic.Strings.Left(Microsoft.VisualBasic.Strings.StrConv(returnValue, vbUnicode), Microsoft.VisualBasic.Strings.Len(GetStr));
        //    //임시주석_E

        //    if (Microsoft.VisualBasic.Strings.Asc(Microsoft.VisualBasic.Strings.Right(returnValue, 1)) == 0)
        //    {
        //        //임시주석_S
        //        //Microsoft.VisualBasic.Strings.Mid(returnValue, Microsoft.VisualBasic.Strings.Len(returnValue), 1) = Microsoft.VisualBasic.Strings.Space(1);
        //        //임시주석_E
        //    }

        //    return returnValue;
        //}
        #endregion

        /// <summary>
        /// 에드온 폼을 운영관리에서 적용한 기본 색으로 바탕색변경
        /// (연말정산 신고용 자료 생성시 폴더 생성, Z, R 클래스에서 사용, 신규 클래스에서는 미사용)
        /// </summary>
        public void Get_FormColor()
        {
            string sQry;
            string StringColor = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select Color from OADM";
                oRecordSet.DoQuery(sQry);

                while (!oRecordSet.EoF)
                {
                    StringColor = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    oRecordSet.MoveNext();
                }

                if (Convert.ToDouble(StringColor) == 1)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5633");
                }
                else if (Convert.ToDouble(StringColor) == 2)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5634");
                }
                else if (Convert.ToDouble(StringColor) == 3)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5635");
                }
                else if (Convert.ToDouble(StringColor) == 4)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5636");
                }
                else if (Convert.ToDouble(StringColor) == 5)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5637");
                }
                else if (Convert.ToDouble(StringColor) == 6)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5638");
                }
                else if (Convert.ToDouble(StringColor) == 7)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5639");
                }
                else if (Convert.ToDouble(StringColor) == 8)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5640");
                }
                else if (Convert.ToDouble(StringColor) == 9)
                {
                    PSH_Globals.SBO_Application.ActivateMenuItem("5641");
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// USER Name 조회
        /// ZPY343 클래스에서만 사용(클래스 미사용으로 구현할 필요 없지만, 만약을 위하여 백업)
        /// </summary>
        /// <param name="oUserSign"></param>
        /// <returns></returns>
        public string Get_UserName(string oUserSign)
        {
            string functionReturnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (!string.IsNullOrEmpty(oUserSign))
                {
                    sQry = "  SELECT U_NAME FROM OUSR";
                    sQry += " WHERE USERID = '" + oUserSign + "'";
                    oRecordSet.DoQuery(sQry);
                    while (!oRecordSet.EoF)
                    {
                        functionReturnValue = oRecordSet.Fields.Item(0).Value;
                        oRecordSet.MoveNext();
                    }

                    if (string.IsNullOrEmpty(functionReturnValue.Trim()))
                    {
                        functionReturnValue = "";
                    }
                }
                else
                {
                    functionReturnValue = "";
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 해당월 마지막 일수 구함
        /// ZPY341 클래스에서만 사용(클래스 미사용으로 구현할 필요 없지만, 만약을 위하여 백업)
        /// </summary>
        /// <param name="JOBDAT"></param>
        /// <returns></returns>
        public string Month_LastDay(string JOBDAT)
        {
            //string functionReturnValue = null;
            
            //switch (true)
            //{
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JOBDAT + "31", "0000-00-00")):
            //        functionReturnValue = Convert.ToString(31);
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JOBDAT + "30", "0000-00-00")):
            //        functionReturnValue = Convert.ToString(30);
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JOBDAT + "29", "0000-00-00")):
            //        functionReturnValue = Convert.ToString(29);
            //        break;
            //    case Information.IsDate(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(JOBDAT + "28", "0000-00-00")):
            //        functionReturnValue = Convert.ToString(28);
            //        break;
            //}
            //return functionReturnValue;

            return DateTime.DaysInMonth(Convert.ToInt16(JOBDAT.Substring(0, 4)), Convert.ToInt16(JOBDAT.Substring(4, 2))).ToString(); //한줄로 끝
        }

        /// <summary>
        /// 테이블 존재 유무와 해당 테이블의 필드명 유무 체크
        /// PH_PY000 클래스에서만 사용(클래스 미사용으로 구현할 필요 없지만, 만약을 위하여 백업)
        /// </summary>
        /// <param name="sTable"></param>
        /// <param name="sField1"></param>
        /// <param name="sField2"></param>
        /// <returns></returns>
        public bool TableFieldCheck(string sTable, string sField1, string sField2)
        {
            bool functionReturnValue;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            functionReturnValue = false;

            sQry = "SELECT * FROM sysobjects WHERE name = '" + sTable + "' AND xtype='U'";
            oRecordSet.DoQuery(sQry);

            if (oRecordSet.RecordCount == 0)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("입력하신 [" + sTable + "테이블이 존재 하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }

            sQry = "select * from INFORMATION_SCHEMA.COLUMNS where table_name='" + sTable + "' and column_name= '" + sField1 + "'";
            oRecordSet.DoQuery(sQry);
            if (oRecordSet.RecordCount == 0)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("입력하신 [" + sField1 + "] 필드명이 존재 하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }

            sQry = "select * from INFORMATION_SCHEMA.COLUMNS where table_name='" + sTable + "' and column_name= '" + sField2 + "'";
            oRecordSet.DoQuery(sQry);
            if (oRecordSet.RecordCount == 0)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("입력하신 [" + sField2 + "] 필드명이 존재 하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }

            functionReturnValue = true;
            return functionReturnValue;
        }

        /// <summary>
        /// 접속자 권한에 따른 아이템 필터
        /// </summary>
        /// <param name="oForm"></param>
        /// <param name="Item">권한을 받는 아이템</param>
        /// <param name="Table">Table Name(ex>@PH_PY001)</param>
        /// <param name="DocType">마스터 : Code, 문서 : DocEntry</param>
        public void AuthorityCheck(SAPbouiCOM.Form oForm, string Item, string Table, string DocType)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "UPDATE [" + Table + "] SET U_NaviDoc = NULL";
                oRecordSet.DoQuery(sQry);

                sQry = "  UPDATE  [" + Table + "]";
                sQry += " SET     U_NaviDoc = " + DocType;
                sQry += " WHERE   U_" + Item + " IN (";
                sQry += "                             SELECT      U_Value";
                sQry += "                             FROM        [@PH_PY000B] T0";
                sQry += "                                         INNER JOIN";
                sQry += "                                         [@PH_PY000A] T1";
                sQry += "                                             ON T0.Code = T1.Code";
                sQry += "                             WHERE       T1.Code = '" + Item + "'";
                sQry += "                                         AND T0.U_UserCode = '" + PSH_Globals.oCompany.UserName + "'";
                sQry += "                             GROUP BY    U_Value";
                sQry += "                           )";

                oRecordSet.DoQuery(sQry);

                oForm.DataBrowser.BrowseBy = "NaviDoc";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 접속자 권한에 따른 사업장 콤보박스 세팅
        /// </summary>
        /// <param name="oForm">화면</param>
        /// <param name="Item"></param>
        /// <param name="AuthorityUse">true:권한에 따른사업장표시, false:전체사업장표시)</param>
        public void CLTCOD_Select(SAPbouiCOM.Form oForm, string Item, bool AuthorityUse)
        {
            int i;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ComboBox oCombo = oForm.Items.Item(Item).Specific;

            try
            {
                if (oCombo.ValidValues.Count > 0)
                {
                    for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }

                if (AuthorityUse == true)
                {
                    sQry = "  SELECT T2.Code ,T2.Name";
                    sQry += " From [@PH_PY000B] T0 INNER JOIN [@PH_PY000A] T1 ON T0.Code = T1.Code";
                    sQry += " INNER JOIN [@PH_PY005A] T2 ON T0.U_Value = T2.Code";
                    sQry += " WHERE T1.Code = 'CLTCOD' AND T0.U_UserCode = '" + PSH_Globals.oCompany.UserName + "'";
                    sQry += " GROUP BY T2.Code ,T2.Name ORDER BY T2.Code";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        while (!oRecordSet.EoF)
                        {
                            oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                            oRecordSet.MoveNext();
                        }
                        
                        oCombo.Select("" + this.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'", "") + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else
                    {
                        oCombo.ValidValues.Add("", "-");
                    }
                } //false
                else
                {
                    sQry = "SELECT Code, Name FROM [@PH_PY005A] ";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        while (!oRecordSet.EoF)
                        {
                            oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                            oRecordSet.MoveNext();
                        }
                    }
                    else
                    {
                        oCombo.ValidValues.Add("", "-");
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
            }
        }

        /// <summary>
        /// FormItemType 반환
        /// </summary>
        /// <param name="pBoFormItemTypesNumber">FormItemType 별 숫자</param>
        /// <returns>FormItemType</returns>
        public BoFormItemTypes ReturnBoFormItemTypesByInteger(short pBoFormItemTypesNumber)
        {
            BoFormItemTypes returnValue = new BoFormItemTypes();

            try
            {
                switch (pBoFormItemTypesNumber)
                {
                    case 102:
                        returnValue = BoFormItemTypes.it_ACTIVE_X;
                        break;
                    case 4:
                        returnValue = BoFormItemTypes.it_BUTTON;
                        break;
                    case 129:
                        returnValue = BoFormItemTypes.it_BUTTON_COMBO;
                        break;
                    case 121:
                        returnValue = BoFormItemTypes.it_CHECK_BOX;
                        break;
                    case 113:
                        returnValue = BoFormItemTypes.it_COMBO_BOX;
                        break;
                    case 16:
                        returnValue = BoFormItemTypes.it_EDIT;
                        break;
                    case 118:
                        returnValue = BoFormItemTypes.it_EXTEDIT;
                        break;
                    case 99:
                        returnValue = BoFormItemTypes.it_FOLDER;
                        break;
                    case 128:
                        returnValue = BoFormItemTypes.it_GRID;
                        break;
                    case 116:
                        returnValue = BoFormItemTypes.it_LINKED_BUTTON;
                        break;
                    case 127:
                        returnValue = BoFormItemTypes.it_MATRIX;
                        break;
                    case 122:
                        returnValue = BoFormItemTypes.it_OPTION_BUTTON;
                        break;
                    case 104:
                        returnValue = BoFormItemTypes.it_PANE_COMBO_BOX;
                        break;
                    case 117:
                        returnValue = BoFormItemTypes.it_PICTURE;
                        break;
                    case 100:
                        returnValue = BoFormItemTypes.it_RECTANGLE;
                        break;
                    case 8:
                        returnValue = BoFormItemTypes.it_STATIC;
                        break;
                    case 131:
                        returnValue = BoFormItemTypes.it_WEB_BROWSER;
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return returnValue;
        }

        /// <summary>
        /// 메트릭스에 컬럼 추가
        /// PH_PY118에서만 사용
        /// </summary>
        /// <param name="oMatrix">메트릭스uid</param>
        /// <param name="Col">컬럼Uid</param>
        /// <param name="iE">컬럼형식-[edit(16),콤보(113), 체크(122), 링크(116)]</param>
        /// <param name="Tn">컬럼타이틀명</param>
        /// <param name="Wt">너비</param>
        /// <param name="Et">Editable true/false값</param>
        /// <param name="St">오른쪽정렬여부</param>
        /// <param name="BouYN">DataBind 여부</param>
        /// <param name="TableNAM">테이블명</param>
        /// <param name="FieldNAM">필드명</param>
        public void PAY_Matrix_AddCol(SAPbouiCOM.Matrix oMatrix, string Col, short iE, string Tn, int Wt, bool Et, bool St, bool BouYN, string TableNAM, string FieldNAM)
        {
            SAPbouiCOM.Columns oCols = null;
            SAPbouiCOM.Column oCol = null;

            try
            {
                oCols = oMatrix.Columns;
                oCol = oCols.Add(Col, ReturnBoFormItemTypesByInteger(iE));

                oCol.DataBind.SetBound(BouYN, TableNAM, FieldNAM); // UI쓸경우 UserDataSources bound먼저해줘야함
                oCol.TitleObject.Caption = Tn;
                oCol.Width = Wt;
                oCol.Editable = Et;
                oCol.RightJustified = St;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCols);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCol);
            }
        }

        /// <summary>
        /// 콤보박스 바인딩
        /// </summary>
        /// <param name="combo">콤보박스 컨트롤명</param>
        /// <param name="query">쿼리</param>
        /// <param name="defaultDescription">기본 선택 Description</param>
        /// <param name="resetWhether">reset 여부</param>
        /// <param name="voidInsertWhether">빈값 추가 여부</param>
        public void Set_ComboList(ComboBox combo, string query, string defaultDescription, bool resetWhether, bool voidInsertWhether)
        {
            SAPbouiCOM.ComboBox ComBox = null;
            SAPbobsCOM.Recordset s_Recordset = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                s_Recordset.DoQuery(query);
                ComBox = combo;

                if (resetWhether == true)
                {
                    while (ComBox.ValidValues.Count > 0)
                    {
                        ComBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }

                if (voidInsertWhether == true)
                {
                    ComBox.ValidValues.Add("", "");
                }

                while (!s_Recordset.EoF)
                {
                    ComBox.ValidValues.Add(s_Recordset.Fields.Item(0).Value.ToString().Trim(), s_Recordset.Fields.Item(1).Value.ToString().Trim());
                    s_Recordset.MoveNext();
                }

                if (!string.IsNullOrEmpty(defaultDescription))
                {
                    ComBox.Select(defaultDescription, SAPbouiCOM.BoSearchKey.psk_ByDescription);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ComBox);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(s_Recordset);
            }
        }

        /// <summary>
        /// 아이디별 사업장 조회
        /// </summary>
        /// <returns>사업장</returns>
        public string User_BPLID()
        {
            string functionReturnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset ooRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select Branch From [OUSR] Where USER_CODE = '" + PSH_Globals.oCompany.UserName.ToString().Trim() + "'";
                ooRecordset01.DoQuery(sQry);

                functionReturnValue = ooRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ooRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 아이디별 사번
        /// </summary>
        /// <returns></returns>
        public string User_MSTCOD()
        {
            string functionReturnValue = null;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select U_MSTCOD From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where b.USER_CODE = '" + PSH_Globals.oCompany.UserName + "'";
                oRecordset01.DoQuery(sQry);

                functionReturnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 아이디별 성명
        /// </summary>
        /// <returns></returns>
        public string User_MSTNAM()
        {
            string functionReturnValue = null;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select U_FULLNAME From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where b.USER_CODE = '" + PSH_Globals.oCompany.UserName + "'";
                oRecordset01.DoQuery(sQry);

                functionReturnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 주민등록번호/외국인등록번호 오류체크
        /// </summary>
        /// <param name="sID">검증하려는 번호</param>
        /// 외국인 주민번호 : 000000-1234567(앞의 6자리는 생년월일,뒷자리는 1:성별구분, 23:등록기관번호, 45:일련번호, 6:등록자구분, 7:검증번호)
        /// <returns></returns>
        public bool GovIDCheck(string sID)
        {
            bool functionReturnValue;
            string Weight;
            int Total;
            int chk;
            int Rmn;
            int i;
            int Dt;
            int Wt;

            functionReturnValue = false;

            sID = sID.Trim();
            if (string.IsNullOrEmpty(sID))
            {
                return functionReturnValue;
            }

            if (sID.Substring(6, 1) == "-")
            {
                sID = sID.Substring(0, 6) + sID.Substring(7, 7);
            }
                
            if (sID.Length != 13)
            {
                return functionReturnValue;
            }

            //성별구분코드(1,2,3,4:내국인, 5,6,7,8:외국인)
            if (Convert.ToInt16(sID.Substring(6, 1)) < 1 || Convert.ToInt16(sID.Substring(6, 1)) > 8)
            {
                return functionReturnValue;
            }
                
            //검증코드
            switch (sID.Substring(6, 1))
            {
                case "5":
                case "6":
                case "7":
                case "8":
                    //외국인
                    //등록기관번호검증
                    if (Convert.ToInt16(sID.Substring(7, 2)) % 2 != 0)
                    {
                        return functionReturnValue;
                    }
                    break;
            }

            chk = Convert.ToInt16(sID.Substring(12, 1));

            Total = 0;
            Weight = "234567892345";

            for (i = 0; i <= 11; i++)
            {
                Dt = Convert.ToInt16(sID.Substring(i, 1));
                Wt = Convert.ToInt16(Weight.Substring(i, 1));

                Total += (Dt * Wt);
            }

            Rmn = 11 - (Total % 11);

            if (Rmn > 9)
            {
                Rmn %= 10;
            }

            switch (sID.Substring(6, 1))
            {
                case "5":
                case "6":
                case "7":
                case "8":
                    // 외국인
                    Rmn += 2;
                    if (Rmn >= 10)
                    {
                        Rmn -= 10;
                    }
                        
                    break;
            }

            functionReturnValue = (Rmn == chk ? true : false);
            return functionReturnValue;
        }

        /// <summary>
        /// 쿼리 실행
        /// </summary>
        /// <param name="sQry">쿼리</param>
        /// <param name="FieldCount">필드위치</param>
        /// <param name="RecordCount">레코드위치</param>
        /// <returns></returns>
        public string GetValue(string sQry, int FieldCount, int RecordCount)
        {
            string functionReturnValue = string.Empty;
            int i;

            SAPbobsCOM.Recordset oRecordset = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordset.DoQuery(sQry);
                if (oRecordset.RecordCount > 0)
                {
                    oRecordset.MoveFirst();
                    if (RecordCount == 0)
                    {
                        RecordCount = 1;
                    }
                    for (i = 1; i <= RecordCount; i++)
                    {
                        functionReturnValue = Convert.ToString(oRecordset.Fields.Item(FieldCount).Value);
                        oRecordset.MoveNext();
                    }
                }
                else
                {
                    functionReturnValue = "";
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 메시지 출력
        /// </summary>
        /// <param name="MDC_pMsg"></param>
        /// <param name="MDC_pType"></param>
        public void MDC_GF_Message(string MDC_pMsg, string MDC_pType)
        {   
            switch (MDC_pType.ToUpper())
            {
                case "S":
                    PSH_Globals.SBO_Application.StatusBar.SetText(MDC_pMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    break;
                case "E":
                    PSH_Globals.SBO_Application.StatusBar.SetText(MDC_pMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    break;
                case "W":
                    PSH_Globals.SBO_Application.StatusBar.SetText(MDC_pMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    break;
            }
        }

        /// <summary>
        /// 사용자정의 Format Search
        /// </summary>
        /// <param name="oForm01"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="ItemUID"></param>
        /// <param name="ColumnUID"></param>
        public void ActiveUserDefineValue(ref SAPbouiCOM.Form oForm01, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string ItemUID, string ColumnUID)
        {
            if (string.IsNullOrEmpty(ColumnUID))
            {
                if (pVal.ItemUID == ItemUID)
                {
                    if (pVal.CharPressed == Convert.ToInt16("9"))
                    {
                        if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
            else
            {
                if (pVal.ItemUID == ItemUID)
                {
                    if (pVal.ColUID == ColumnUID)
                    {
                        if (pVal.CharPressed == Convert.ToInt16("9"))
                        {
                            if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Columns(ColumnUID).Cells(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 사용자정의 Format Search #2(적용된 TextBox에 값이 있어도 무조건 호출)
        /// </summary>
        /// <param name="oForm01"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="ItemUID"></param>
        /// <param name="ColumnUID"></param>
        public void ActiveUserDefineValueAlways(ref SAPbouiCOM.Form oForm01, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string ItemUID, string ColumnUID)
        {
            if (string.IsNullOrEmpty(ColumnUID))
            {
                if (pVal.ItemUID == ItemUID)
                {
                    if (pVal.CharPressed == Convert.ToInt16("9"))
                    {
                        if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                    }
                }
            }
            else
            {
                if (pVal.ItemUID == ItemUID)
                {
                    if (pVal.ColUID == ColumnUID)
                    {
                        if (pVal.CharPressed == Convert.ToDouble("9"))
                        {
                            if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Columns(ColumnUID).Cells(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// String Format의 일자를 Date Format 일자로 변경
        /// </summary>
        /// <param name="pDate">String Format의 일자</param>
        /// <param name="pChar">구분 Character</param>
        /// <returns>Date Format 일자</returns>
        public string ConvertDateType(string pDate, string pChar)
        {
            string returnValue = pDate.Substring(0, 4) + pChar + pDate.Substring(4, 2) + pChar + pDate.Substring(6, 2);
            return returnValue;
        }

        #region MDC_PS_Common 클래스 메소드 구현_S

        /// <summary>
        /// 콤보박스 Value Insert
        /// </summary>
        /// <param name="pFormUID">FormID</param>
        /// <param name="pItemUID">ItemID</param>
        /// <param name="pColumnUID">ColumnID</param>
        /// <param name="pVALUE">Value</param>
        /// <param name="pDescription">Description</param>
        public void Combo_ValidValues_Insert(string pFormUID, string pItemUID, string pColumnUID, string pVALUE, string pDescription)
        {
            try
            {
                this.DoQuery("EXEC COMBO_VALIDVALUES_INSERT '" + pFormUID + "','" + pItemUID + "','" + pColumnUID + "','" + pVALUE + "','" + pDescription + "'");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combo_ValidValues_SetValueColumn
        /// </summary>
        /// <param name="pColumn"></param>
        /// <param name="pFormUID"></param>
        /// <param name="pItemUID"></param>
        /// <param name="pColumnUID"></param>
        /// <param name="pEmptyValue"></param>
        public void Combo_ValidValues_SetValueColumn(SAPbouiCOM.Column pColumn, string pFormUID, string pItemUID, string pColumnUID, bool pEmptyValue)
        {
            int loopCount;
            string Query01;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT VALUE, DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" + pFormUID + "' AND ITEMUID = '" + pItemUID + "' AND COLUMNUID = '" + pColumnUID + "'";
                oRecordSet.DoQuery(Query01);

                if (oRecordSet.RecordCount > 0)
                {
                    for (loopCount = 1; loopCount <= pColumn.ValidValues.Count; loopCount++)
                    {
                        pColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    if (pEmptyValue == true)
                    {
                        pColumn.ValidValues.Add("", "");
                    }
                    for (loopCount = 1; loopCount <= oRecordSet.RecordCount; loopCount++)
                    {
                        pColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 쿼리 실행
        /// </summary>
        /// <param name="pQuery01">쿼리</param>
        public void DoQuery(string pQuery01)
        {
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery(pQuery01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// EnableMenu 설정(VB6.0함수명:MDC_GP_EnableMenus)
        /// </summary>
        /// <param name="pForm">Form객체</param>
        /// <param name="pPreview">인쇄[미리보기]</param>
        /// <param name="pPrint">인쇄[출력]</param>
        /// <param name="pDeleteRow">행삭제</param>
        /// <param name="pFind">문서찾기</param>
        /// <param name="pAdd">문서추가</param>
        /// <param name="pNextRecord">다음</param>
        /// <param name="pPreviousRecord">이전</param>
        /// <param name="pFirstRecord">맨처음</param>
        /// <param name="pLastRecord">맨끝</param>
        /// <param name="pCancel">문서취소</param>
        /// <param name="pRowAdd">행추가</param>
        /// <param name="pDuplicate">문서복제</param>
        /// <param name="pRemove">문서제거</param>
        /// <param name="pRowClose">행닫기</param>
        /// <param name="pClose">문서닫기</param>
        /// <param name="pRestore">문서복원</param>
        public void SetEnableMenus(SAPbouiCOM.Form pForm, bool pPreview, bool pPrint, bool pDeleteRow, bool pFind, bool pAdd, bool pNextRecord, bool pPreviousRecord, bool pFirstRecord, bool pLastRecord, bool pCancel, bool pRowAdd, bool pDuplicate, bool pRemove, bool pRowClose, bool pClose, bool pRestore)
        {
            try
            {
                pForm.EnableMenu("519", pPreview); // 인쇄[미리보기]
                pForm.EnableMenu("520", pPrint); // 인쇄[출력]
                pForm.EnableMenu("1281", pFind); //문서찾기
                pForm.EnableMenu("1282", pAdd); //문서추가
                pForm.EnableMenu("1283", pRemove); //문서제거(데이터 삭제시 사용)
                pForm.EnableMenu("1284", pCancel); //문서취소
                pForm.EnableMenu("1285", pRestore); //문서복원
                pForm.EnableMenu("1286", pClose); //문서닫기
                pForm.EnableMenu("1287", pDuplicate); //문서복제
                pForm.EnableMenu("1288", pNextRecord); //다음
                pForm.EnableMenu("1289", pPreviousRecord); //이전
                pForm.EnableMenu("1290", pFirstRecord); //맨처음
                pForm.EnableMenu("1291", pLastRecord); //맨끝
                pForm.EnableMenu("1292", pRowAdd); //행추가
                pForm.EnableMenu("1293", pDeleteRow); //행삭제
                pForm.EnableMenu("1299", pRowClose); //행닫기
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combo_ValidValues_SetValueItem
        /// </summary>
        /// <param name="Combo"></param>
        /// <param name="FormUID"></param>
        /// <param name="ItemUID"></param>
        /// <param name="EmptyValue"></param>
        public void Combo_ValidValues_SetValueItem(SAPbouiCOM.ComboBox Combo, string FormUID, string ItemUID, bool EmptyValue)
        {
            int i;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT VALUE,DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" + FormUID + "' AND ITEMUID = '" + ItemUID + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount > 0)
                {
                    for (i = 1; i <= Combo.ValidValues.Count; i++)
                    {
                        Combo.ValidValues.Remove((0));
                    }
                    if (EmptyValue == true)
                    {
                        Combo.ValidValues.Add("", "");
                    }
                    for (i = 1; i <= oRecordset01.RecordCount; i++)
                    {
                        Combo.ValidValues.Add(oRecordset01.Fields.Item(0).Value, oRecordset01.Fields.Item(1).Value);
                        oRecordset01.MoveNext();
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
        }

        /// <summary>
        /// 품목 단중 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_UnWeight(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_UnWeight FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 대분류 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_ItmBsort(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 판매기준단위 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_SbasUnit(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_SBasUnit FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 매입기준단위 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_ObasUnit(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_OBasUnit FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 단위수량1 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_Unit1(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_UnitQ1 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }
            
        /// <summary>
        /// 품목 규격1 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_Spec1(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_Spec1 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 규격2 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_Spec2(string ItemCode)
        {
            string functionReturnValue = null;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_Spec2 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 규격3 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_Spec3(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_Spec3 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 관리 기준 조회(배치관리품목)
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_ManBtchNum(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT ManBtchNum FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 품목 거래형태 조회
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <returns></returns>
        public string GetItem_TradeType(string ItemCode)
        {
            string functionReturnValue = string.Empty;
            string Query01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "SELECT U_TradeType FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
                oRecordset01.DoQuery(Query01);

                if (oRecordset01.RecordCount == 0)
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordset01.Fields.Item(0).Value;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 선행작업의 총중량 - 현재 작업에서 생성된 중량을 제외한 값을 구함
        /// </summary>
        /// <param name="oForm01">Form</param>
        public void SBO_SetBackOrderFunction(ref SAPbouiCOM.Form oForm01)
        {
            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                return;
            }

            string BaseType;
            string BaseTable;
            int BaseEntry;
            int BaseLine;

            string errCode = string.Empty;

            int i;
            string Query01;

            SAPbouiCOM.Matrix oMat01 = oForm01.Items.Item("38").Specific;
            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                if (oMat01.VisualRowCount > 1)
                {
                    
                    for (i = 1; i <= oMat01.RowCount - 1; i++)
                    {
                        BaseType = oMat01.Columns.Item("43").Cells.Item(i).Specific.Value;

                        if (BaseType == "-1") //BaseType이 "-1"이면 예외처리
                        {
                            errCode = "1";
                            throw new Exception();
                        }

                        BaseEntry = oMat01.Columns.Item("45").Cells.Item(i).Specific.Value;
                        BaseLine = oMat01.Columns.Item("46").Cells.Item(i).Specific.Value;
                        
                        if (BaseType == "17") ////판매오더
                        {
                            BaseTable = "RDR";
                        }
                        else if (BaseType == "23") //판매견적
                        {
                            BaseTable = "QUT";
                        }
                        else if (BaseType == "15") //납품
                        {
                            BaseTable = "DLN";
                        }
                        else if (BaseType == "16") //판매반품
                        {
                            BaseTable = "RDN";
                        }
                        else if (BaseType == "13") //AR송장
                        {
                            BaseTable = "INV";
                        }
                        else if (BaseType == "14") //AR대변메모
                        {
                            BaseTable = "RIN";
                            
                        }
                        else if (BaseType == "22") //구매오더
                        {
                            BaseTable = "POR";
                        }
                        else if (BaseType == "20") //입고PO
                        {
                            BaseTable = "PDN";
                        }
                        else if (BaseType == "21") //구매반품
                        {
                            BaseTable = "RPD";
                        }
                        else if (BaseType == "18") //AP송장
                        {
                            BaseTable = "PCH";
                        }
                        else if (BaseType == "19") //AP대변메모
                        {
                            BaseTable = "RPC";
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.MessageBox("화면캡쳐후 관리자에게 문의바랍니다.");
                            return;
                        }

                        Query01 = " PS_SBO_GETQUANTITY '" + BaseType + "','" + BaseTable + "','" + BaseEntry + "','" + BaseLine + "'";
                        oRecordset01.DoQuery(Query01);

                        oMat01.Columns.Item("U_Qty").Cells.Item(i).Specific.Value = System.Math.Round(oRecordset01.Fields.Item(0).Value, 0);
                        oMat01.Columns.Item("11").Cells.Item(i).Specific.Value = System.Math.Round(oRecordset01.Fields.Item(1).Value, 2);
                        oMat01.Columns.Item("1").Cells.Item(oMat01.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
            }
            catch(Exception ex)
            {
                if(errCode == "1") //BaseType이 "-1"이면 아무것도 안함
                {
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
        }

        /// <summary>
        /// ItemName에 작은 따옴표가 있을경우 하나 더 추가
        /// </summary>
        /// <param name="ItemName"></param>
        /// <returns></returns>
        public string Make_ItemName(string ItemName)
        {
            string returnValue = string.Empty;
            int i;
            string tempItemName = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                for (i = 0; i < ItemName.Length; i++)
                {
                    tempItemName += codeHelpClass.Mid(ItemName, i, 1);
                    if (codeHelpClass.Mid(ItemName, i, 1) == "'")
                    {
                        tempItemName += "'";
                    }
                }

                returnValue = tempItemName.Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 아이디별 창고 선택 [기본창고 1, 외주가공 8, 임가공 9]
        /// </summary>
        /// <param name="Gbn"></param>
        /// <returns></returns>
        public string User_WhsCode(string Gbn)
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  a.WhsCode";
                sQry += " FROM    [OWHS] a";
                sQry += "         INNER JOIN";
                sQry += "         [OUSR] b";
                sQry += "             ON a.BPLid = b.Branch";
                sQry += " WHERE   b.USER_CODE = '" + PSH_Globals.oCompany.UserName.Trim() + "'";
                sQry += "         AND LEFT(WhsCode, 1) = '" + Gbn + "'";
                sQry += "         AND RIGHT(a.WhsCode, 2) = RIGHT(b.DfltsGroup, 2)";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 부서코드 조회
        /// </summary>
        /// <returns>DeptCode</returns>
        public string User_DeptCode()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  dept";
                sQry += " FROM    [OHEM] a";
                sQry += "         INNER JOIN";
                sQry += "         [OUSR] b";
                sQry += "             ON a.userId = b.USERID";
                sQry += " WHERE   USER_CODE = '" + PSH_Globals.oCompany.UserName.Trim() + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 팀코드 조회
        /// </summary>
        /// <returns>DeptCode</returns>
        public string User_TeamCode()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select U_TeamCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '"+ PSH_Globals.oCompany.UserName.Trim() + "'";
                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 담당코드 조회
        /// </summary>
        /// <returns>RspCode</returns>
        public string User_RspCode()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  U_RspCode";
                sQry += " FROM    [OHEM] a";
                sQry += "         INNER JOIN";
                sQry += "         [OUSR] b";
                sQry += "             ON a.userId = b.USERID";
                sQry += " WHERE   USER_CODE = '" + PSH_Globals.oCompany.UserName.Trim() + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 반코드 조회
        /// </summary>
        /// <returns>ClsCode</returns>
        public string User_ClsCode()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  U_ClsCode";
                sQry += " FROM    [OHEM] a";
                sQry += "         INNER JOIN";
                sQry += "         [OUSR] b";
                sQry += "             ON a.userId = b.USERID";
                sQry += " WHERE   USER_CODE = '" + PSH_Globals.oCompany.UserName.Trim() + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 SuperUser 여부
        /// </summary>
        /// <returns> Y:수퍼유저, N:일반유저</returns>
        public string User_SuperUserYN()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  T0.SUPERUSER";
                sQry += " FROM    OUSR AS T0";
                sQry += " WHERE   T0.User_Code = '" + PSH_Globals.oCompany.UserName.Trim() + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 입력일자가 현재 서버일자보다 미래가 될 수 없도록 제한
        /// </summary>
        /// <param name="inputdate">미래일자 여부</param>
        /// <returns></returns>
        public string Future_Date_Check(string inputdate)
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT    CASE ";
                sQry += "               WHEN CONVERT(CHAR(8), GETDATE(), 112) >= '" + inputdate.Trim() + "' THEN 'Y'";
                sQry += "               ELSE 'N'";
                sQry += "           END";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 접속한 사용자의 주요업무 조회
        /// </summary>
        /// <returns>주요업무(인사마스터(OHEM)의 Remark)</returns>
        public string User_MainJob()
        {
            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT    T0.Remark";
                sQry += " FROM      OHEM AS T0";
                sQry += "           LEFT JOIN";
                sQry += "           OUSR AS T1";
                sQry += "               ON T0.UserID = T1.USERID";
                sQry += " WHERE     T1.User_Code = '" + PSH_Globals.oCompany.UserName.Trim() + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item(0).Value.ToString().Trim();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 중량계산
        /// </summary>
        /// <param name="ItemCode">품목코드</param>
        /// <param name="Qty">수량</param>
        /// <param name="BPLId">사업장</param>
        /// <returns></returns>
        public double Calculate_Weight(string ItemCode, int Qty, string BPLId)
        {
            double returnValue = 0;
            string sQry;
            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT    U_OBasUnit,";
                sQry += "           U_UnitQ1,";
                sQry += "           U_Spec1,";
                sQry += "           U_Spec2,";
                sQry += "           U_Spec3,";
                sQry += "           U_UnWeight";
                sQry += " FROM      [OITM]";
                sQry += " WHERE     ItemCode = '" + ItemCode + "'";

                oRecordset01.DoQuery(sQry);

                if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "101")
                {
                    returnValue = Qty;
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "102")
                {
                    returnValue = Qty * Convert.ToDouble(oRecordset01.Fields.Item(1).Value.ToString().Trim());
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "201")
                {
                    returnValue = (Convert.ToDouble(oRecordset01.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString())) * Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim()) * 0.02808 * (Convert.ToDouble(oRecordset01.Fields.Item(4).Value.ToString().Trim()) / 1000) * Qty;
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "202")
                {
                    returnValue = Qty * Convert.ToDouble(oRecordset01.Fields.Item(5).Value.ToString().Trim()) / 1000;
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "203")
                {
                    returnValue = 0;
                }

                if (BPLId == "3" || BPLId == "5")
                {
                    returnValue = System.Math.Round(returnValue, 2);
                }
                else
                {
                    returnValue = System.Math.Round(returnValue, 0);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ItemCode"></param>
        /// <param name="Weight"></param>
        /// <returns></returns>
        public int Calculate_Qty(string ItemCode, int Weight)
        {
            int returnValue = 0;
            double tempReturnValue = 0;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT    U_OBasUnit,";
                sQry += "           U_UnitQ1,";
                sQry += "           U_Spec1,";
                sQry += "           U_Spec2,";
                sQry += "           U_Spec3,";
                sQry += "           U_UnWeight";
                sQry += " FROM      [OITM]";
                sQry += " WHERE     ItemCode = '" + ItemCode + "'";

                oRecordset01.DoQuery(sQry);

                if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "101")
                {
                    tempReturnValue = Weight;
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "102")
                {
                    if (string.IsNullOrEmpty(oRecordset01.Fields.Item(1).Value.ToString().Trim()) || Convert.ToDouble(oRecordset01.Fields.Item(1).Value.ToString().Trim()) == 0)
                    {
                        tempReturnValue = 0;
                    }
                    else
                    {
                        tempReturnValue = Weight / Convert.ToDouble(oRecordset01.Fields.Item(1).Value.ToString().Trim());
                    }
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "201")
                {
                    if ((Convert.ToDouble(oRecordset01.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim())) * Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim()) * 0.02808 * (Convert.ToDouble(oRecordset01.Fields.Item(4).Value.ToString().Trim()) / 1000) == Convert.ToDouble("") || (Convert.ToDouble(oRecordset01.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim())) * Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim()) * 0.02808 * (Convert.ToDouble(oRecordset01.Fields.Item(4).Value.ToString().Trim()) / 1000) == 0)
                    {
                        tempReturnValue = 0;
                    }
                    else
                    {
                        tempReturnValue = Weight / ((Convert.ToDouble(oRecordset01.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim())) * Convert.ToDouble(oRecordset01.Fields.Item(3).Value.ToString().Trim()) * 0.02808 * (Convert.ToDouble(oRecordset01.Fields.Item(4).Value.ToString().Trim()) / 1000));
                    }
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "202")
                {
                    if (string.IsNullOrEmpty(oRecordset01.Fields.Item(5).Value.ToString().Trim()) || Convert.ToDouble(oRecordset01.Fields.Item(5).Value.ToString().Trim()) == 0)
                    {
                        tempReturnValue = 0;
                    }
                    else
                    {
                        tempReturnValue = Weight / Convert.ToDouble(oRecordset01.Fields.Item(5).Value.ToString().Trim()) * 1000;
                    }
                }
                else if (oRecordset01.Fields.Item(0).Value.ToString().Trim() == "203")
                {
                    tempReturnValue = 0;
                }

                returnValue = Convert.ToInt32(System.Math.Round(tempReturnValue, 0));
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// R3 통신(실제 구현 할때 테스트 및 보완 필요, 2020.09.15 송명규)
        /// </summary>
        /// <param name="BPLId"></param>
        /// <param name="ItemCode"></param>
        /// <param name="ItemName"></param>
        /// <param name="Size"></param>
        /// <param name="Qty"></param>
        /// <param name="Unit"></param>
        /// <param name="RequestDate"></param>
        /// <param name="DueDate"></param>
        /// <param name="ItemType"></param>
        /// <param name="RequestNo"></param>
        /// <returns></returns>
        public string RFC_Sender(string BPLId, string ItemCode, string ItemName, string Size, double Qty, string Unit, string RequestDate, string DueDate, string ItemType, string RequestNo)
        {
            string returnValue = string.Empty;
            string tempReturnValue;
            string WERKS;
            string errCode = string.Empty;
            string errMessage = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            RfcConfigParameters rfc = new RfcConfigParameters();
            
            //if (i == 0)
            //{
            //rfc.Add(RfcConfigParameters.Name, "");
            rfc.Add(RfcConfigParameters.AppServerHost, "192.1.1.217");  //서버IP
            //rfc.Add(RfcConfigParameters.AppServerHost, "192.1.11.7");  //테스트
            rfc.Add(RfcConfigParameters.SystemNumber, "00");
            rfc.Add(RfcConfigParameters.User, "ifuser");   //사용자ID
            rfc.Add(RfcConfigParameters.Password, "pdauser");  //사용자 패스워드
            rfc.Add(RfcConfigParameters.Client, "210");
            rfc.Add(RfcConfigParameters.Language, "KO");
            rfc.Add(RfcConfigParameters.PoolSize, "5");

            //rfc.Add(RfcConfigParameters.MaxPoolSize, "10");
            //rfc.Add(RfcConfigParameters.IdleTimeout, "500");
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
            RfcRepository rfcRep = rfcDest.Repository;

            //PSH_Globals.oSapConnection01 = Interaction.CreateObject("SAP.Functions");
            //PSH_Globals.oSapConnection01.Connection.User = "ifuser";
            //PSH_Globals.oSapConnection01.Connection.Password = "pdauser";
            ////oSapConnection01.Connection.client = "710"
            //PSH_Globals.oSapConnection01.Connection.Client = "210";
            ////oSapConnection01.Connection.ApplicationServer = "192.1.11.7"
            //PSH_Globals.oSapConnection01.Connection.ApplicationServer = "192.1.1.217";
            //PSH_Globals.oSapConnection01.Connection.Language = "KO";
            //PSH_Globals.oSapConnection01.Connection.SystemNumber = "00";

            //if (!PSH_Globals.oSapConnection01.Connection.Logon(0, true))
            //{
            //    errCode = "1";
            //    throw new Exception();
            //}

            IRfcFunction oFunction01 = rfcRep.CreateFunction("ZMM_INTF_GROUP");
            //}

            try
            {
                if (Convert.ToDouble(BPLId) == 1)
                {
                    WERKS = "9200";
                }
                else if (Convert.ToDouble(BPLId) == 2)
                {
                    WERKS = "9300";
                }
                else
                {
                    WERKS = "9200";
                }

                oFunction01.SetValue("I_WERKS", WERKS); //플랜트 홀딩스 창원 9200, 홀딩스 부산 9300
                oFunction01.SetValue("I_MATNR", ItemCode); //자재코드 char(18)
                oFunction01.SetValue("I_MAKTX", ItemName); //자재내역 char(40)
                oFunction01.SetValue("I_WRKST", Size); //자재규격 char(48)
                oFunction01.SetValue("I_MENGE", Qty); //구매요청수량 dec(13,3)
                oFunction01.SetValue("I_MEINS", Unit); //단위 char(3)
                oFunction01.SetValue("I_BADAT", RequestDate); //구매요청일 char(8)
                oFunction01.SetValue("I_LFDAT", DueDate); //납품일 char(8)
                oFunction01.SetValue("I_MATKL", ItemType); //자재그룹 char(9)
                oFunction01.SetValue("I_ZBANFN", RequestNo); //구매요청번호

                oFunction01.Invoke(rfcDest);
                
                errMessage = oFunction01.GetValue("E_MESSAGE").ToString();

                if (string.IsNullOrEmpty(errMessage)) //에러메시지
                {
                    tempReturnValue = oFunction01.GetString("E_BANFN").ToString() + "/" + oFunction01.GetString("E_BNFPO").ToString(); //통합구매요청번호, 통합구매요청 품목번호
                }
                else
                {
                    errCode = "1";
                    throw new Exception();
                    //MDC_Com.MDC_GF_Message(ref oFunction01.Imports("E_MESSAGE").Value, ref "E");
                    //goto RFC_Sender_Exit;
                }

                //oFunction01.Exports("I_WERKS") = WERKS; //플랜트 홀딩스 창원 9200, 홀딩스 부산 9300
                //oFunction01.Exports("I_MATNR") = ItemCode; //자재코드 char(18)
                //oFunction01.Exports("I_MAKTX") = ItemName; //자재내역 char(40)
                //oFunction01.Exports("I_WRKST") = Size; //자재규격 char(48)
                //oFunction01.Exports("I_MENGE") = Qty; //구매요청수량 dec(13,3)
                //oFunction01.Exports("I_MEINS") = Unit; //단위 char(3)
                //oFunction01.Exports("I_BADAT") = RequestDate; //구매요청일 char(8)
                //oFunction01.Exports("I_LFDAT") = DueDate; //납품일 char(8)
                //oFunction01.Exports("I_MATKL") = ItemType; //자재그룹 char(9)
                //oFunction01.Exports("I_ZBANFN") = RequestNo; //구매요청번호

                //if (!(oFunction01.Call))
                //{
                //    errCode = "2";
                //    throw new Exception();
                //    //MDC_Com.MDC_GF_Message(ref "안강(R/3)서버 함수호출중 오류발생", ref "E");
                //    //goto RFC_Sender_Exit;
                //}
                //else
                //{
                //    if ((string.IsNullOrEmpty(oFunction01.Imports("E_MESSAGE").Value))) //에러메시지
                //    {
                //        tempReturnValue = oFunction01.Imports("E_BANFN").Value + "/" + oFunction01.Imports("E_BNFPO").Value; ////통합구매요청번호 '//통합구매요청 품목번호
                //    }
                //    else
                //    {
                //        errCode = "3";
                //        throw new Exception();
                //        //MDC_Com.MDC_GF_Message(ref oFunction01.Imports("E_MESSAGE").Value, ref "E");
                //        //goto RFC_Sender_Exit;
                //    }
                //}

                returnValue = tempReturnValue;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    dataHelpClass.MDC_GF_Message(errMessage, "E");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                //if ((PSH_Globals.oSapConnection01.Connection != null))
                //{
                //    if (i == LastRow)
                //    {
                //        PSH_Globals.oSapConnection01.Connection.Logoff();
                //        PSH_Globals.oSapConnection01 = null;
                //    }
                //}

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFunction01);
            }

            return returnValue;
        }

        /// <summary>
        /// KPI 평가등급 계산
        /// </summary>
        /// <param name="prmBaseEntry">KPI목표문서번호</param>
        /// <param name="prmBaseLine">KPI목표문서행번호</param>
        /// <param name="prmTableName">KPI목표 테이블 명</param>
        /// <param name="prmResult">실적</param>
        /// <param name="prmMonth">실적등록월</param>
        /// <returns>KPI평가등급</returns>
        public string Cal_KPI_Grade(short prmBaseEntry, short prmBaseLine, string prmTableName, string prmResult, string prmMonth)
        {
            //1. 해당KPI목표 테이블의 문서번호와 행번호를 이용하여 A~E까지의 값 조회
            //2. 등급기준(최대, 최소)에 따라 분기문이 달라져야 하므로 등급기준이 최대인지, 최소인지 함께 조회

            string returnValue = string.Empty;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PS_Z_GetKPIGrade " + prmBaseEntry + "," + prmBaseLine + ",'" + prmTableName + "','" + prmResult + "', '" + prmMonth + "'";

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item("Grade").Value;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        /// <summary>
        /// KPI 평가점수 계산
        /// </summary>
        /// <param name="prmKPIGrade">KPI평가등급</param>
        /// <returns>KPI평가점수</returns>
        public double Cal_KPI_Score(string prmKPIGrade)
        {
            double returnValue = 0;
            string sQry;
            double KPI_Score = 0;
            short loopCount01;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  T1.U_CodeNm AS [CodeName],";
                sQry += "         T1.U_Num1 AS [KPIScore]";
                sQry += " FROM    [@PS_HR200H] AS T0";
                sQry += "         INNER JOIN";
                sQry += "         [@PS_HR200L] AS T1";
                sQry += "             ON T0.Code = T1.Code";
                sQry += " WHERE   T0.Name = '평가점수'";

                oRecordset01.DoQuery(sQry);

                for (loopCount01 = 0; loopCount01 <= oRecordset01.RecordCount - 1; loopCount01++)
                {
                    if (prmKPIGrade == oRecordset01.Fields.Item("CodeName").Value)
                    {
                        KPI_Score = oRecordset01.Fields.Item("KPIScore").Value;
                    }

                    oRecordset01.MoveNext();
                }

                returnValue = KPI_Score;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
            
            return returnValue;
        }

        /// <summary>
        /// KPI 진척률(달성률) 조회
        /// </summary>
        /// <param name="prmBasEntry">목표문서번호</param>
        /// <param name="prmBasLine">목표행번호</param>
        /// <param name="prmDocType">문서타입</param>
        /// <param name="prmRslt">실적</param>
        /// <returns>KPI 진척률</returns>
        public double Cal_KPI_AchieveRate(short prmBasEntry, short prmBasLine, string prmDocType, string prmRslt)
        {
            double returnValue = 0;
            string sQry;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PS_Z_GetAchieveRate " + prmBasEntry + "," + prmBasLine + ",'" + prmDocType + "','" + prmRslt + "'"; //진척률 계산 프로시져

                oRecordset01.DoQuery(sQry);

                returnValue = oRecordset01.Fields.Item("AchieveRate").Value;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 마감상태 조회
        /// </summary>
        /// <param name="prmBPLId">사업장</param>
        /// <param name="prmDocDate">등록일</param>
        /// <param name="prmFormTypeEx">화면타입(UID)</param>
        /// <returns></returns>
        public bool Check_Finish_Status(string prmBPLId, string prmDocDate, object prmFormTypeEx)
        {
            bool returnValue = false;
            string sQry;
            string CheckFinishStatus;

            SAPbobsCOM.Recordset oRecordset01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = " EXEC PS_Z_CheckFinishStatus '";
                sQry += prmBPLId + "','";
                sQry += prmDocDate + "','";
                sQry += prmFormTypeEx + "'";

                oRecordset01.DoQuery(sQry);

                CheckFinishStatus = oRecordset01.Fields.Item("ReturnValue").Value;

                if (CheckFinishStatus == "True")
                {
                    returnValue = true;
                }
                else
                {
                    returnValue = false;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset01);
            }

            return returnValue;
        }

        #endregion MDC_PS_Common 클래스 메소드 구현_E

        /// <summary>
        /// SAPConnection
        /// </summary>
        /// <param name="rfcDest">목적지 IP</param>
        /// <param name="rfcRep">저장소Class</param>
        /// <param name="pName"></param>
        /// <param name="pAppServerHost">Server IP</param>
        /// <param name="pClient">Client No</param>
        /// <param name="pUser">ID</param>
        /// <param name="pPassword">Password</param>
        /// <returns></returns>
        public bool SAPConnection(ref RfcDestination rfcDest, ref RfcRepository rfcRep, string pName, string pAppServerHost, string pClient, string pUser, string pPassword)
        {
            bool returnValue;
            //RfcDestination rfcDest;
            //RfcRepository rfcRep;

            try
            {
                rfcDest = this.RfcDestination(pName, pAppServerHost, pClient, pUser, pPassword);
                rfcRep = rfcDest.Repository;

                returnValue = true;
            }
            catch
            {
                returnValue = false;
            }

            return returnValue;
        }

        /// <summary>
        /// FTP FIleConnection , Upload
        /// </summary>
        /// <param name="ip"></param>
        /// <param name="port"></param>
        /// <param name="userid"></param>
        /// <param name="pwd"></param>
        /// <param name="path">업로드할 파일 저장할 FTP경로</param>
        /// <param name="filename">업로드 대상 파일 경로</param>
        /// <returns></returns>
        public bool FTPConn_Upload(string ip, string port, string userid, string pwd, string path, string filename)
        {
            bool isConnected = false;
            int contentLen;
            int buffLength = 2048;
            byte[] buff = new byte[buffLength]; //버퍼사이즈 지정
            FileInfo fileInfo = new FileInfo(filename);
            string uri = string.Format(@"FTP://{0}:{1}/{2}/{3}", ip, port, path, fileInfo.Name);
            FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(new Uri(uri));//FtpWebRequest object 생성
            ftpRequest.Credentials = new NetworkCredential(userid, pwd);  //아이디, 패스워드 검증
            ftpRequest.KeepAlive = false; //서버에 대한 연결이 소멸되지 않아야 하면 true, 소멸되어야 하면 false  (KeepAlive의 기본값은 원래 true임.)
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile; //지정한 업로드 명령을 실행
            ftpRequest.UsePassive = true; //passive모드 사용여부
            ftpRequest.UseBinary = true; //전송 타입설정
            ftpRequest.ContentLength = fileInfo.Length;  //서버에 파일사이즈를 알림

            try
            {
                FileStream fs = fileInfo.OpenRead(); //파일 읽기
                
                Stream strm = ftpRequest.GetRequestStream(); //업로드 할 파일 스트림을 가져옴.
                contentLen = fs.Read(buff, 0, buffLength); //2kb씩 파일 스트림을 읽은 후 길이 반환
                
                while (contentLen != 0) //스트림을 다 읽을때까지 반복.
                {
                    strm.Write(buff, 0, contentLen);//FTP에 파일을 기록
                    contentLen = fs.Read(buff, 0, buffLength);
                }
                strm.Close();
                fs.Close();

                isConnected = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            return isConnected;
        }

        /// <summary>
        /// SAP R3 RFC Destination 구현
        /// </summary>
        /// <param name="pName">Server Name</param>
        /// <param name="pAppServerHost">Server IP</param>
        /// <param name="pClient">Client</param>
        /// <param name="pUser">UserID</param>
        /// <param name="pPassword">UserPassword</param>
        /// <returns></returns>
        public RfcDestination RfcDestination(string pName, string pAppServerHost, string pClient, string pUser, string pPassword)
        {
            RfcConfigParameters rfc = new RfcConfigParameters();

            rfc.Add(RfcConfigParameters.Name, pName);
            rfc.Add(RfcConfigParameters.AppServerHost, pAppServerHost);
            rfc.Add(RfcConfigParameters.Client, pClient);
            rfc.Add(RfcConfigParameters.User, pUser);
            rfc.Add(RfcConfigParameters.Password, pPassword);

            RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);

            return rfcDest;
        }

        /// <summary>
        /// 본사 R3 서버 정보 조회
        /// </summary>
        /// <returns>Array[0] : Client, Array[1] : ServerIP</returns>
        public string[] GetR3ServerInfo()
        {
            string[] r3ServerInfo = new string[2];

            if (GetValue("SELECT DB_NAME()", 0, 1) == "PSHDB") //운영용DB
            {
                r3ServerInfo[0] = GetValue("SELECT U_Client FROM [@PSH_R3INFO] WHERE Code = '1'", 0, 1); //Client(210)
                r3ServerInfo[1] = GetValue("SELECT U_ServerIP FROM [@PSH_R3INFO] WHERE Code = '1'", 0, 1); //ServerIP(192.1.11.3)
            }
            else //그 외(테스트DB)
            {
                r3ServerInfo[0] = GetValue("SELECT U_Client FROM [@PSH_R3INFO] WHERE Code = '2'", 0, 1); //Client(810)
                r3ServerInfo[1] = GetValue("SELECT U_ServerIP FROM [@PSH_R3INFO] WHERE Code = '2'", 0, 1); //ServerIP(192.1.11.7)
            }

            return r3ServerInfo;
        }
    }
}
