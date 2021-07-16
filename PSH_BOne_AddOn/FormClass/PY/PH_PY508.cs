using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 재직증명 등록 및 발급
    /// </summary>
    internal class PH_PY508 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY508A; //헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY508B; //라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY508.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY508_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY508");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry"

                oForm.Freeze(true);

                PH_PY508_CreateItems();
                PH_PY508_FormItemEnabled();
                PH_PY508_ComboBox_Setting(); //콤보박스 세팅
                PH_PY508_EnableMenus();
                PH_PY508_SetDocument(oFormDocEntry);
                PH_PY508_FormResize();
                PH_PY508_LoadCaption();
                PH_PY508_FormReset(); //폼초기화 추가(2013.01.29 송명규)

                oMat01.Columns.Item("CLTAdrs").Visible = false; //사업장주소 Visible False
                oMat01.Columns.Item("RepName").Visible = false; //대표이사명 Visible False

                oForm.Items.Item("SFrDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                oForm.Items.Item("SToDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.ActiveItem = "MSTCOD";
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY508_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY508A = oForm.DataSources.DBDataSources.Item("@PH_PY508A");
                oDS_PH_PY508B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //메트릭스 개체
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                ////////////기본정보//////////_S
                //기본정보는 실제 테이블과 필드 매핑
                ////////////기본정보//////////_E

                ////////////조회정보//////////_S
                //관리번호
                oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

                //사업장
                oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

                //문서번호1
                oForm.DataSources.UserDataSources.Add("SDocNo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("SDocNo1").Specific.DataBind.SetBound(true, "", "SDocNo1");

                //문서번호2
                oForm.DataSources.UserDataSources.Add("SDocNo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("SDocNo2").Specific.DataBind.SetBound(true, "", "SDocNo2");

                //발행일(FR)
                oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

                //발행일(TO)
                oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

                //사번
                oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("SMSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SMSTNAM").Specific.DataBind.SetBound(true, "", "SMSTNAM");

                //소속팀
                oForm.DataSources.UserDataSources.Add("STeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STeamCode").Specific.DataBind.SetBound(true, "", "STeamCode");

                //소속담당
                oForm.DataSources.UserDataSources.Add("SRspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SRspCode").Specific.DataBind.SetBound(true, "", "SRspCode");

                //용도(%)
                oForm.DataSources.UserDataSources.Add("SUseCmt", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
                oForm.Items.Item("SUseCmt").Specific.DataBind.SetBound(true, "", "SUseCmt");

                //용도 타입
                oForm.DataSources.UserDataSources.Add("UseType1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 10);
                oForm.Items.Item("UseType1").Specific.DataBind.SetBound(true, "", "UseType1");
                ////////////조회정보//////////_E
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY508_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY508_ComboBox_Setting()
        {
            string sQry;
            string CLTCOD;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

                ////////////기본정보//////////_S
                //소속팀
                oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = '1'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
                oForm.Items.Item("TeamCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //소속담당
                oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = '2'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
                oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //직책
                oForm.Items.Item("Position").Specific.ValidValues.Add("%", "선택");
                sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Position").Specific, sQry, "", false, false);
                oForm.Items.Item("Position").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //용도 타입
                //20190322 황영수 수정S
                sQry = "select u_code ,u_codenm  from [@PS_HR200L] where code ='88' order by u_num1";
                dataHelpClass.Set_ComboList(oForm.Items.Item("UseType1").Specific, sQry, "", false, false);
                oForm.Items.Item("UseType1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //20190322 황영수 수정E
                ////////////기본정보//////////_E

                ////////////매트릭스//////////_S
                //사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");

                //직책
                sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Position"), sQry, "", "");
                ////////////매트릭스//////////_E
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY508_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY508_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY508_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY508_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY860_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY508_FormClear()
        {
            int DocEntry;
            string sQry;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY508A]";
                RecordSet01.DoQuery(sQry);

                DocEntry = Convert.ToInt32(RecordSet01.Fields.Item(0).Value.ToString().Trim());

                if (DocEntry == 0)
                {
                    oDS_PH_PY508A.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PH_PY508A.SetValue("DocEntry", 0, Convert.ToString(DocEntry + 1));
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY508_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_FormResize_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PH_PY508_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    oForm.Items.Item("BtnDelete").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    oForm.Items.Item("BtnDelete").Enabled = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_LoadCaption_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PH_PY508_FormReset()
        {
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                PH_PY508_FormClear(); //DocEntry 초기화

                ////////////기준정보//////////
                oDS_PH_PY508A.SetValue("U_CLTCOD", 0, dataHelpClass.User_BPLID()); //사업장
                oDS_PH_PY508A.SetValue("U_DocNo1", 0, ""); //문서번호1
                oDS_PH_PY508A.SetValue("U_DocNo2", 0, ""); //문서번호2
                oDS_PH_PY508A.SetValue("U_CLTAdrs", 0, ""); //사업장주소
                oDS_PH_PY508A.SetValue("U_RepName", 0, ""); //대표이사명
                oDS_PH_PY508A.SetValue("U_RegDt", 0, DateTime.Now.ToString("yyyyMMdd")); //발행일자
                oDS_PH_PY508A.SetValue("U_MSTCOD", 0, ""); //사번
                oDS_PH_PY508A.SetValue("U_MSTNAM", 0, ""); //성명
                oDS_PH_PY508A.SetValue("U_TeamCode", 0, "%"); //소속팀
                oDS_PH_PY508A.SetValue("U_RspCode", 0, "%"); //소속담당
                oDS_PH_PY508A.SetValue("U_Position", 0, "%"); //직책
                oDS_PH_PY508A.SetValue("U_GovID", 0, ""); //주민등록번호
                oDS_PH_PY508A.SetValue("U_CurAdrs", 0, ""); //현주소
                oDS_PH_PY508A.SetValue("U_GrpDat", 0, DateTime.Now.ToString("yyyyMMdd")); //입사일
                oDS_PH_PY508A.SetValue("U_UseCmt", 0, ""); //용도
                oDS_PH_PY508A.SetValue("U_Retire", 0, "N"); //퇴사자

                PH_PY508_GetDocNo(); //문서번호생성

                oForm.Items.Item("MSTCOD").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 문서번호 생성
        /// </summary>
        private void PH_PY508_GetDocNo()
        {
            string sQry;
            string StdYear = string.Empty;
            string CLTCOD;
            string regDt;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                regDt = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim();
                if (regDt != "") //LoadForm이 실행될때 FormItemEnabled의 콤보박스 Select 이벤트가 발생하여 발행일자(regDt)의 값이 할당되지 않은 상태에서 Substring을 할 경우 오류 발생, 따라서 발행일자(regDt)에 값이 할당되었을 때만 Substring 실시
                {
                    StdYear = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim().Substring(0, 4);
                }

                sQry = "EXEC PH_PY508_05 '" + CLTCOD + "', '" + StdYear + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("DocNo1").Specific.Value = StdYear;
                oForm.Items.Item("DocNo2").Specific.Value = oRecordSet01.Fields.Item("DocNo2").Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_GetDocNo_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        private void PH_PY508_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PH_PY508B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY508B.Offset = oRow;
                oDS_PH_PY508B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_Add_MatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {

            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY508_MTX01()
        {
            short i;
            short errNum = 0;
            string sQry;

            string sDocEntry; //관리번호
            string sCLTCOD; //사업장
            string SDocNo1; //문서번호1
            string SDocNo2; //문서번호2
            string SFrDate; //발행일(시작)
            string SToDate; //발행일(종료)
            string sMSTCOD; //사번
            string sTeamCode; //소속팀
            string sRspCode; //소속담당
            string SUseCmt; //용도

            sDocEntry = oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim(); //관리번호
            sCLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim(); //사업장
            SDocNo1 = oForm.Items.Item("SDocNo1").Specific.Value.ToString().Trim(); //문서번호1
            SDocNo2 = oForm.Items.Item("SDocNo2").Specific.Value.ToString().Trim(); //문서번호2
            SFrDate = oForm.Items.Item("SFrDate").Specific.Value.ToString().Trim(); //발행일(시작)
            SToDate = oForm.Items.Item("SToDate").Specific.Value.ToString().Trim(); //발행일(종료)
            sMSTCOD = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim(); //사번
            sTeamCode = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim(); //소속팀
            sRspCode = oForm.Items.Item("SRspCode").Specific.Value.ToString().Trim(); //소속담당
            SUseCmt = oForm.Items.Item("SUseCmt").Specific.Value.ToString().Trim(); //용도

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                sQry = "EXEC [PH_PY508_01] '";
                sQry += sDocEntry + "','"; //관리번호
                sQry += sCLTCOD + "','"; //사업장
                sQry += SDocNo1 + "','"; //문서번호1
                sQry += SDocNo2 + "','"; //문서번호2
                sQry += SFrDate + "','"; //발행일(시작)
                sQry += SToDate + "','"; //발행일(종료)
                sQry += sMSTCOD + "','"; //사원번호
                sQry += sTeamCode + "','"; //소속팀
                sQry += sRspCode + "','"; //소속담당
                sQry += SUseCmt + "'"; //용도

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY508B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errNum = 1;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY508_LoadCaption();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY508B.Size)
                    {
                        oDS_PH_PY508B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY508B.Offset = i;

                    oDS_PH_PY508B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY508B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim()); //선택
                    oDS_PH_PY508B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
                    oDS_PH_PY508B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim()); //사업장
                    oDS_PH_PY508B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("DocNo1").Value.ToString().Trim()); //문서번호1
                    oDS_PH_PY508B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("DocNo2").Value.ToString().Trim()); //문서번호2
                    oDS_PH_PY508B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("CLTAdrs").Value.ToString().Trim()); //사업장주소
                    oDS_PH_PY508B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("RepName").Value.ToString().Trim()); //대표이사명
                    oDS_PH_PY508B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("RegDt").Value.ToString().Trim()); //발행일
                    oDS_PH_PY508B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //사번
                    oDS_PH_PY508B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim()); //성명
                    oDS_PH_PY508B.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim()); //소속팀
                    oDS_PH_PY508B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim()); //소속팀명
                    oDS_PH_PY508B.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim()); //소속담당
                    oDS_PH_PY508B.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("RspName").Value.ToString().Trim()); //소속담당명
                    oDS_PH_PY508B.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("Position").Value.ToString().Trim()); //직책
                    oDS_PH_PY508B.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("GovID").Value.ToString().Trim()); //주민등록번호
                    oDS_PH_PY508B.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("CurAdrs").Value.ToString().Trim()); //현주소
                    oDS_PH_PY508B.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("GrpDat").Value.ToString().Trim()); //입사일
                    oDS_PH_PY508B.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("UseCmt").Value.ToString().Trim()); //용도
                    oDS_PH_PY508B.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("Retire").Value.ToString().Trim()); //퇴사여부

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 기본정보 삭제
        /// </summary>
        private void PH_PY508_DeleteData()
        {
            short errNum = 0;
            string sQry;
            string DocEntry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PH_PY508A] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.RecordCount == 0)
                    {
                        errNum = 1;
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "EXEC PH_PY508_04 '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("삭제대상이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_DeleteData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 기본정보 수정
        /// </summary>
        /// <returns>성공:true, 실패:false</returns>
        private bool PH_PY508_UpdateData()
        {
            bool functionReturnValue = false;
            string sQry;

            int DocEntry; //관리번호
            string DocNo1; //문서번호1
            string DocNo2; //문서번호2
            string CLTCOD; //사업장
            string CLTAdrs; //사업장주소
            string RepName; //대표이사명
            string RegDt; //발행일
            string MSTCOD; //사번
            string MSTNAM; //성명
            string TeamCode; //소속팀(코드)
            string RspCode; //소속담당(코드)
            string Position; //직책
            string GovID; //주민등록번호
            string CurAdrs; //현주소
            string GrpDat; //입사일
            string UseCmt; //용도
            string Retire; //퇴사여부

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = Convert.ToInt16(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()); //관리번호
                DocNo1 = oForm.Items.Item("DocNo1").Specific.Value.ToString().Trim(); //문서번호1
                DocNo2 = oForm.Items.Item("DocNo2").Specific.Value.ToString().Trim(); //문서번호2
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                CLTAdrs = oForm.Items.Item("CLTAdrs").Specific.Value.ToString().Trim(); //사업장주소
                RepName = oForm.Items.Item("RepName").Specific.Value.ToString().Trim(); //대표이사명
                RegDt = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim(); //발행일
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //사번
                MSTNAM = oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim(); //성명
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //소속팀(코드)
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim(); //소속담당(코드)
                Position = oForm.Items.Item("Position").Specific.Value.ToString().Trim(); //직책
                GovID = oForm.Items.Item("GovID").Specific.Value.ToString().Trim(); //주민등록번호
                CurAdrs = oForm.Items.Item("CurAdrs").Specific.Value.ToString().Trim(); //현주소
                GrpDat = oForm.Items.Item("GrpDat").Specific.Value.ToString().Trim(); //입사일
                UseCmt = oForm.Items.Item("UseCmt").Specific.Value.ToString().Trim(); //용도
                Retire = (oForm.Items.Item("Retire").Specific.Checked == true ? "1" : "0"); //퇴사여부

                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
                }

                sQry = "EXEC [PH_PY508_03] '";
                sQry += DocEntry + "','"; //관리번호
                sQry += DocNo1 + "','"; //문서번호1
                sQry += DocNo2 + "','"; //문서번호2
                sQry += CLTCOD + "','"; //사업장
                sQry += CLTAdrs + "','"; //사업장주소
                sQry += RepName + "','"; //대표이사명
                sQry += RegDt + "','"; //발행일
                sQry += MSTCOD + "','"; //사원번호
                sQry += MSTNAM + "','"; //사원성명
                sQry += TeamCode + "','"; //소속팀(코드)
                sQry += RspCode + "','"; //소속담당(코드)
                sQry += Position + "','"; //직책
                sQry += GovID + "','"; //주민등록번호
                sQry += CurAdrs + "','"; //현주소
                sQry += GrpDat + "','"; //입사일
                sQry += UseCmt + "','"; //용도
                sQry += Retire + "'"; //퇴사여부

                RecordSet01.DoQuery(sQry);

                functionReturnValue = true;
                PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_DeleteData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 INSERT
        /// </summary>
        /// <returns>성공:true, 실패:false</returns>
        private bool PH_PY508_AddData()
        {
            bool functionReturnValue = false;
            string sQry;

            int DocEntry; //관리번호
            string DocNo1; //문서번호1
            string DocNo2; //문서번호2
            string CLTCOD; //사업장
            string CLTAdrs; //사업장주소
            string RepName; //대표이사명
            string RegDt; //발행일
            string MSTCOD; //사번
            string MSTNAM; //성명
            string TeamCode; //소속팀(코드)
            string RspCode; //소속담당(코드)
            string Position; //직책
            string GovID; //주민등록번호
            string CurAdrs; //현주소
            string GrpDat; //입사일
            string UseCmt; //용도
            string Retire; //퇴사여부
            string UserSign; //UserSign

            DocNo1 = oForm.Items.Item("DocNo1").Specific.Value.ToString().Trim(); //문서번호1
            DocNo2 = oForm.Items.Item("DocNo2").Specific.Value.ToString().Trim(); //문서번호2
            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
            CLTAdrs = oForm.Items.Item("CLTAdrs").Specific.Value.ToString().Trim(); //사업장주소
            RepName = oForm.Items.Item("RepName").Specific.Value.ToString().Trim(); //대표이사명
            RegDt = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim(); //발행일
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //사번
            MSTNAM = oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim(); //성명
            TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //소속팀(코드)
            RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim(); //소속담당(코드)
            Position = oForm.Items.Item("Position").Specific.Value.ToString().Trim(); //직책
            GovID = oForm.Items.Item("GovID").Specific.Value.ToString().Trim(); //주민등록번호
            CurAdrs = oForm.Items.Item("CurAdrs").Specific.Value.ToString().Trim(); //현주소
            GrpDat = oForm.Items.Item("GrpDat").Specific.Value.ToString().Trim(); //입사일
            UseCmt = oForm.Items.Item("UseCmt").Specific.Value.ToString().Trim(); //용도
            Retire = (oForm.Items.Item("Retire").Specific.Checked == true ? "1" : "0"); //퇴사여부
            UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY508A]";
                RecordSet01.DoQuery(sQry);

                if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(RecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }

                sQry = "EXEC [PH_PY508_02] '";
                sQry += DocEntry + "','"; //관리번호
                sQry += DocNo1 + "','"; //문서번호1
                sQry += DocNo2 + "','"; //문서번호2
                sQry += CLTCOD + "','"; //사업장
                sQry += CLTAdrs + "','"; //사업장주소
                sQry += RepName + "','"; //대표이사명
                sQry += RegDt + "','"; //발행일
                sQry += MSTCOD + "','"; //사원번호
                sQry += MSTNAM + "','"; //사원성명
                sQry += TeamCode + "','"; //소속팀(코드)
                sQry += RspCode + "','"; //소속담당(코드)
                sQry += Position + "','"; //직책
                sQry += GovID + "','"; //주민등록번호
                sQry += CurAdrs + "','"; //현주소
                sQry += GrpDat + "','"; //입사일
                sQry += UseCmt + "','"; //용도
                sQry += Retire + "','"; //퇴사여부
                sQry += UserSign + "'"; //UserSign

                RecordSet02.DoQuery(sQry);

                functionReturnValue = true;
                PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_AddData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PH_PY508_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short errNum = 0;
            
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("DocNo1").Specific.Value.ToString().Trim())) //문서번호
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("DocNo2").Specific.Value.ToString().Trim())) //문서번호2
                {
                    errNum = 2;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("CLTAdrs").Specific.Value.ToString().Trim())) //사업장주소
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("RepName").Specific.Value.ToString().Trim()))//대표이사명
                {
                    errNum = 4;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("RegDt").Specific.Value.ToString().Trim())) //발행일
                {
                    errNum = 5;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim())) //사번
                {
                    errNum = 6;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("GovID").Specific.Value.ToString().Trim())) //주민등록번호
                {
                    errNum = 7;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("CurAdrs").Specific.Value.ToString().Trim())) //현주소
                {
                    errNum = 8;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("GrpDat").Specific.Value.ToString().Trim())) //입사일
                {
                    errNum = 9;
                    throw new Exception();
                }
                else if(string.IsNullOrEmpty(oForm.Items.Item("UseCmt").Specific.Value.ToString().Trim())) //용도
                {
                    errNum = 10;
                    throw new Exception();
                }

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("문서번호1은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DocNo1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("문서번호2는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DocNo2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장주소는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTAdrs").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("대표이사명은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("발행일은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("RegDt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("주민등록번호는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("GovID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 8)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("현주소는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CurAdrs").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 9)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입사일은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("GrpDat").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 10)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("용도는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("UseCmt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_HeaderSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        /// <param name="pFormMode"></param>
        private void PH_PY508_FlushToItemValue(string oUID, int oRow, string oCol, string pFormMode)
        {
            int loopCount;
            string sQry;
            string TeamCode;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {

                    case "RegDt":

                        if (pFormMode != "2") //업데이트 모드에서는 문서번호 재설정 안함
                        {
                            PH_PY508_GetDocNo();
                        }
                        break;

                    case "CLTCOD":

                        PH_PY508_GetDocNo();
                        break;

                    case "SCLTCOD":

                        if (oForm.Items.Item("STeamCode").Specific.ValidValues.Count > 0)
                        {
                            for (loopCount = oForm.Items.Item("STeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                            {
                                oForm.Items.Item("STeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        oForm.Items.Item("STeamCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT    U_Code,";
                        sQry += "           U_CodeNm";
                        sQry += " FROM      [@PS_HR200L]";
                        sQry += " WHERE     Code = '1'";
                        sQry += "           AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "           AND U_UseYN = 'Y'";
                        sQry += " ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("STeamCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "STeamCode":

                        TeamCode = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("SRspCode").Specific.ValidValues.Count > 0)
                        {
                            for (loopCount = oForm.Items.Item("SRspCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                            {
                                oForm.Items.Item("SRspCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //담당콤보세팅
                        oForm.Items.Item("SRspCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT      U_Code AS [Code],";
                        sQry += "             U_CodeNm As [Name]";
                        sQry += " FROM        [@PS_HR200L]";
                        sQry += " WHERE       Code = '2'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += "             AND U_Char1 = '" + TeamCode + "'";
                        sQry += " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("SRspCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("SRspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "MSTCOD":

                        sQry = "EXEC PH_PY508_06 '";
                        sQry += oForm.Items.Item("CLTCOD").Specific.Value + "','";
                        sQry += oForm.Items.Item("MSTCOD").Specific.Value + "'";

                        oRecordSet01.DoQuery(sQry);

                        oDS_PH_PY508A.SetValue("U_MSTNAM", 0, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim()); //성명
                        oDS_PH_PY508A.SetValue("U_TeamCode", 0, oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim()); //팀코드
                        oDS_PH_PY508A.SetValue("U_RspCode", 0, oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim()); //소속담당
                        oDS_PH_PY508A.SetValue("U_Position", 0, oRecordSet01.Fields.Item("Position").Value.ToString().Trim()); //직책
                        oDS_PH_PY508A.SetValue("U_GovID", 0, oRecordSet01.Fields.Item("GovID").Value.ToString().Trim()); //주민등록번호
                        oDS_PH_PY508A.SetValue("U_CurAdrs", 0, oRecordSet01.Fields.Item("CurAdrs").Value.ToString().Trim()); //현주소
                        oDS_PH_PY508A.SetValue("U_GrpDat", 0, oRecordSet01.Fields.Item("GrpDat").Value.ToString("yyyyMMdd")); //입사일
                        oDS_PH_PY508A.SetValue("U_CLTAdrs", 0, oRecordSet01.Fields.Item("CLTAdrs").Value.ToString().Trim()); //사업장주소
                        oDS_PH_PY508A.SetValue("U_RepName", 0, oRecordSet01.Fields.Item("RepName").Value.ToString().Trim()); //대표이사명

                        oForm.Items.Item("Retire").Enabled = false;
                        break;

                    case "SMSTCOD":

                        oForm.Items.Item("SMSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.Value + "'", ""); //성명
                        break;

                    case "UseType1":
                        if (oForm.Items.Item("UseType1").Specific.Selected.Value == "M")
                        {
                            oForm.Items.Item("UseCmt").Specific.Value = "";
                        }
                        else
                        {
                            oForm.Items.Item("UseCmt").Specific.Value = dataHelpClass.Get_ReData("u_codenm", "u_code", "[@PS_HR200L]", "'" + oForm.Items.Item("UseType1").Specific.Selected.Value + "'", " and code ='88'");
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PH_PY508_Print_Report01()
        {
            string WinTitle;
            string ReportName = string.Empty;
            string sQry;
            short loopCount;
            string CLTCOD;
            string MSTCOD;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                MSTCOD = dataHelpClass.User_MSTCOD(); //조회자사번

                //임시테이블에 check된항목저장하기 위한 기 정보 삭제
                sQry = "DELETE [Z_PH_PY508] WHERE MSTCOD = '" + MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);

                //임시테이블에 check된 항목저장
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY508B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        sQry = "INSERT INTO [Z_PH_PY508] VALUES ('" + MSTCOD + "', '" + oDS_PH_PY508B.GetValue("U_ColReg02", loopCount).ToString().Trim() + "')";
                        oRecordSet01.DoQuery(sQry);
                    }
                }

                WinTitle = "[PH_PY508] 재직증명원";
                
                if (CLTCOD == "1" || CLTCOD == "3") //창원, 포장
                {
                    ReportName = "PH_PY508_01.rpt";
                } //부산
                else if (CLTCOD == "2")
                {
                    ReportName = "PH_PY508_02.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD)); //사번

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY508_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "BtnAdd") //추가/확인
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY508_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PH_PY508_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY508_FormReset();

                            oForm.Items.Item("Retire").Enabled = true;

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY508_LoadCaption();
                            PH_PY508_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                            if (PH_PY508_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PH_PY508_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY508_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY508_LoadCaption();
                            PH_PY508_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {

                        PH_PY508_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        
                        PH_PY508_LoadCaption();
                        PH_PY508_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {

                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            PH_PY508_DeleteData();
                            PH_PY508_FormReset();

                            oForm.Items.Item("Retire").Enabled = true;

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY508_LoadCaption();
                            PH_PY508_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY508_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", ""); //기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SMSTCOD", ""); //조회조건-사번
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY508_FlushToItemValue(pVal.ItemUID, 0, "", "");
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);

                            

                            //DataSource를 이용하여 각 컨트롤에 값을 출력
                            oDS_PH_PY508A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //관리번호
                            oDS_PH_PY508A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //사업장
                            oDS_PH_PY508A.SetValue("U_DocNo1", 0, oMat01.Columns.Item("DocNo1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //문서번호1
                            oDS_PH_PY508A.SetValue("U_DocNo2", 0, oMat01.Columns.Item("DocNo2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //문서번호2
                            oDS_PH_PY508A.SetValue("U_CLTAdrs", 0, oMat01.Columns.Item("CLTAdrs").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //사업장주소
                            oDS_PH_PY508A.SetValue("U_RepName", 0, oMat01.Columns.Item("RepName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //대표이사명
                            oDS_PH_PY508A.SetValue("U_RegDt", 0, oMat01.Columns.Item("RegDt").Cells.Item(pVal.Row).Specific.Value.ToString().Replace(".", "")); //발행일
                            oDS_PH_PY508A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //사번
                            oDS_PH_PY508A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //성명
                            oDS_PH_PY508A.SetValue("U_TeamCode", 0, oMat01.Columns.Item("TeamCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //소속팀
                            oDS_PH_PY508A.SetValue("U_RspCode", 0, oMat01.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //소속담당
                            oDS_PH_PY508A.SetValue("U_Position", 0, oMat01.Columns.Item("Position").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //직책
                            oDS_PH_PY508A.SetValue("U_GovID", 0, oMat01.Columns.Item("GovID").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //주민등록번호
                            oDS_PH_PY508A.SetValue("U_CurAdrs", 0, oMat01.Columns.Item("CurAdrs").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //현주소
                            oDS_PH_PY508A.SetValue("U_GrpDat", 0, oMat01.Columns.Item("GrpDat").Cells.Item(pVal.Row).Specific.Value.ToString().Replace(".", "")); //입사일
                            oDS_PH_PY508A.SetValue("U_UseCmt", 0, oMat01.Columns.Item("UseCmt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //용도
                            oDS_PH_PY508A.SetValue("U_Retire", 0, (oMat01.Columns.Item("Retire").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "1" ? "Y" : "N")); //퇴사여부

                            oForm.Items.Item("Retire").Enabled = false;

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PH_PY508_LoadCaption();
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            PH_PY508_FlushToItemValue(pVal.ItemUID, 0, "", Convert.ToString(pVal.FormMode));
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY508_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY508A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY508B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY508_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_ROW_DELETE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        {
            int i;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PH_PY508A.RemoveRecord(oDS_PH_PY508A.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY508_Add_MatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY508A.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY508_Add_MatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            PH_PY508_FormReset(); 
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PH_PY508_LoadCaption();
                            oForm.Items.Item("Retire").Enabled = true;
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;

                        case "7169": //엑셀 내보내기, 엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            oForm.Freeze(true);
                            PH_PY508_Add_MatrixRow(oMat01.VisualRowCount, false);
                            oForm.Freeze(false);
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기 처리
                            oForm.Freeze(true);
                            oDS_PH_PY508B.RemoveRecord(oDS_PH_PY508B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            break;
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
