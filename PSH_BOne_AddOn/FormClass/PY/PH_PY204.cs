using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 교육계획등록
    /// </summary>
    internal class PH_PY204 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY204B; //등록라인
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY204.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY204_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY204");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry"

                oForm.Freeze(true);
                PH_PY204_CreateItems();
                PH_PY204_ComboBox_Setting();
                PH_PY204_EnableMenus();
                PH_PY204_SetDocument(oFormDocEntry);
                PH_PY204_FormResize();
                PH_PY204_Add_MatrixRow(0, true);
                PH_PY204_LoadCaption();
                PH_PY204_FormReset();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY204_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY204B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //재직여부
                oForm.DataSources.UserDataSources.Add("Status", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Status").Specific.DataBind.SetBound(true, "", "Status");

                //사원번호
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //사원성명
                oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");

                //기준월
                oForm.DataSources.UserDataSources.Add("StdMt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdMt").Specific.DataBind.SetBound(true, "", "StdMt");

                //계획구분
                oForm.DataSources.UserDataSources.Add("PlnCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PlnCls").Specific.DataBind.SetBound(true, "", "PlnCls");

                //상태
                oForm.DataSources.UserDataSources.Add("DocStatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DocStatus").Specific.DataBind.SetBound(true, "", "DocStatus");

                //교육명
                oForm.DataSources.UserDataSources.Add("EduName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("EduName").Specific.DataBind.SetBound(true, "", "EduName");

                //교육기관
                oForm.DataSources.UserDataSources.Add("EduOrg", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("EduOrg").Specific.DataBind.SetBound(true, "", "EduOrg");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY204_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                //재직여부
                oForm.Items.Item("Status").Specific.ValidValues.Add("1", "재직자");
                oForm.Items.Item("Status").Specific.ValidValues.Add("2", "퇴직자포함");
                oForm.Items.Item("Status").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //기준월
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm AS [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = '4'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";

                oForm.Items.Item("StdMt").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("StdMt").Specific, sQry, "", false, false);
                oForm.Items.Item("StdMt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //계획구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm AS [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P239'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";

                oForm.Items.Item("PlnCls").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("PlnCls").Specific, sQry, "", false, false);
                oForm.Items.Item("PlnCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //상태
                oForm.Items.Item("DocStatus").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("DocStatus").Specific.ValidValues.Add("O", "계획");
                oForm.Items.Item("DocStatus").Specific.ValidValues.Add("P", "완료");
                oForm.Items.Item("DocStatus").Specific.ValidValues.Add("C", "취소");
                oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                ////////////매트릭스//////////_S
                //상태
                oMat01.Columns.Item("DocStatus").ValidValues.Add("O", "계획");
                oMat01.Columns.Item("DocStatus").ValidValues.Add("P", "완료");
                oMat01.Columns.Item("DocStatus").ValidValues.Add("C", "취소");

                //계획구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm AS [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P239'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("PlnCls"), sQry, "", "");

                //이수(수료)증
                oMat01.Columns.Item("Certi").ValidValues.Add("O", "미제출");
                oMat01.Columns.Item("Certi").ValidValues.Add("C", "제출");

                //보고서
                oMat01.Columns.Item("Report").ValidValues.Add("O", "미제출");
                oMat01.Columns.Item("Report").ValidValues.Add("C", "제출");

                //사업장
                sQry = "  SELECT      BPLId AS [BPLId],";
                sQry += "             BPLName AS [BPLName]";
                sQry += " FROM        OBPL";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), sQry, "", "");
                ////////////매트릭스//////////_E
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY204_EnableMenus()
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY204_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY204_FormItemEnabled();
                    ////Call PH_PY021_AddMatrixRow(0, True) '
                }
                else
                {
                    //        oForm.Mode = fm_FIND_MODE
                    //        Call PH_PY204_FormItemEnabled
                    //        oForm.Items("DocEntry").Specific.Value = oFormDocEntry
                    //        oForm.Items("1").Click ct_Regular
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY204_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY204_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_FormResize_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY204_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PH_PY204B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY204B.Offset = oRow;
                oDS_PH_PY204B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PH_PY204B.SetValue("U_ColReg01", oRow, "Y");

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_Add_MatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PH_PY204_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_LoadCaption_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PH_PY204_FormReset()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                string User_BPLID = dataHelpClass.User_BPLID();

                //헤더 초기화
                oForm.DataSources.UserDataSources.Item("CLTCOD").Value = User_BPLID; //사업장
                oForm.DataSources.UserDataSources.Item("TeamCode").Value = "%"; //부서
                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = ""; //사번
                oForm.DataSources.UserDataSources.Item("MSTNAM").Value = ""; //성명
                oForm.DataSources.UserDataSources.Item("Status").Value = "1"; //재직여부
                oForm.DataSources.UserDataSources.Item("StdYear").Value = DateTime.Now.ToString("yyyy"); //기준년도
                oForm.DataSources.UserDataSources.Item("DocStatus").Value = "%"; //상태
                oForm.DataSources.UserDataSources.Item("PlnCls").Value = "%"; //계획구분
                oForm.DataSources.UserDataSources.Item("EduName").Value = ""; //교육명
                oForm.DataSources.UserDataSources.Item("EduOrg").Value = ""; //교육기관

                //라인 초기화
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                PH_PY204_Add_MatrixRow(0, true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY204_MTX01()
        {
            int i;
            string sQry;
            short ErrNum = 0;

            string CLTCOD; //사업장
            string TeamCode; //부서
            string MSTCOD; //사번
            string StdYear; //기준년도
            string StdMt; //기준월
            string Status; //재직여부
            string DocStatus; //상태
            string PlnCls; //계획구분
            string EduName; //교육명
            string EduOrg; //교육기관

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();  //사업장
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //부서
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //사번
                StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim(); //기준년도
                StdMt = oForm.Items.Item("StdMt").Specific.Value.ToString().Trim(); //기준월
                Status = oForm.Items.Item("Status").Specific.Value.ToString().Trim(); //재직여부
                DocStatus = oForm.Items.Item("DocStatus").Specific.Value.ToString().Trim(); //상태
                PlnCls = oForm.Items.Item("PlnCls").Specific.Value.ToString().Trim(); //계획구분
                EduName = oForm.Items.Item("EduName").Specific.Value.ToString().Trim(); //교육명
                EduOrg = oForm.Items.Item("EduOrg").Specific.Value.ToString().Trim(); //교육기관

                sQry = "EXEC [PH_PY204_01] '";
                sQry += CLTCOD + "','"; //사업장
                sQry += TeamCode + "','"; //부서
                sQry += MSTCOD + "','"; //사번
                sQry += StdYear + "','"; //기준년도
                sQry += StdMt + "','"; //기준월
                sQry += Status + "','"; //재직여부
                sQry += DocStatus + "','"; //상태
                sQry += PlnCls + "','"; //계획구분
                sQry += EduName + "','"; //교육명
                sQry += EduOrg + "'"; //교육기관

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY204B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    PH_PY204_Add_MatrixRow(0, true);
                    PH_PY204_LoadCaption();

                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY204B.Size)
                    {
                        oDS_PH_PY204B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY204B.Offset = i;

                    oDS_PH_PY204B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY204B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim()); //선택
                    oDS_PH_PY204B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("IDX").Value.ToString().Trim()); //IDX
                    oDS_PH_PY204B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //대상자사번
                    oDS_PH_PY204B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim()); //대상자성명
                    oDS_PH_PY204B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("StdYear").Value.ToString().Trim()); //기준년도
                    oDS_PH_PY204B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim()); //부서코드
                    oDS_PH_PY204B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim()); //부서명
                    oDS_PH_PY204B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim()); //담당코드
                    oDS_PH_PY204B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("RspName").Value.ToString().Trim()); //담당명
                    oDS_PH_PY204B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("ClsCode").Value.ToString().Trim()); //반코드
                    oDS_PH_PY204B.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("ClsName").Value.ToString().Trim()); //반명
                    oDS_PH_PY204B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("JIGNAM").Value.ToString().Trim()); //직급
                    oDS_PH_PY204B.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("Class").Value.ToString().Trim()); //구분
                    oDS_PH_PY204B.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("ClassName").Value.ToString().Trim()); //구분명
                    oDS_PH_PY204B.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("EduName").Value.ToString().Trim()); //교육명
                    oDS_PH_PY204B.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("EduOrg").Value.ToString().Trim()); //교육기관
                    oDS_PH_PY204B.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("EduQuarter").Value.ToString().Trim()); //분기
                    oDS_PH_PY204B.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("EduHarf").Value.ToString().Trim()); //반기
                    oDS_PH_PY204B.SetValue("U_ColSum02", i, oRecordSet01.Fields.Item("EduMonth").Value.ToString().Trim()); //월
                    oDS_PH_PY204B.SetValue("U_ColSum03", i, oRecordSet01.Fields.Item("EduAmount").Value.ToString().Trim()); //교육비
                    oDS_PH_PY204B.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("EduHour").Value.ToString().Trim()); //교육시간
                    oDS_PH_PY204B.SetValue("U_ColDt01", i, oRecordSet01.Fields.Item("EduFrDt").Value.ToString("yyyyMMdd")); //교육시작일
                    oDS_PH_PY204B.SetValue("U_ColDt02", i, oRecordSet01.Fields.Item("EduToDt").Value.ToString("yyyyMMdd")); //교육종료일
                    oDS_PH_PY204B.SetValue("U_ColReg24", i, oRecordSet01.Fields.Item("Reference").Value.ToString().Trim()); //관련근거
                    oDS_PH_PY204B.SetValue("U_ColReg25", i, oRecordSet01.Fields.Item("Comment").Value.ToString().Trim()); //비고
                    oDS_PH_PY204B.SetValue("U_ColReg26", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim()); //사업장
                    oDS_PH_PY204B.SetValue("U_ColReg27", i, oRecordSet01.Fields.Item("DocStatus").Value.ToString().Trim()); //상태
                    oDS_PH_PY204B.SetValue("U_ColReg28", i, oRecordSet01.Fields.Item("PlnCls").Value.ToString().Trim()); //계획구분
                    oDS_PH_PY204B.SetValue("U_ColReg29", i, oRecordSet01.Fields.Item("Certi").Value.ToString().Trim()); //이수증(수료증)
                    oDS_PH_PY204B.SetValue("U_ColReg30", i, oRecordSet01.Fields.Item("Report").Value.ToString().Trim()); //보고서
                    oDS_PH_PY204B.SetValue("U_ColSum04", i, oRecordSet01.Fields.Item("EduPlnCnt").Value.ToString().Trim()); //건수
                    oDS_PH_PY204B.SetValue("U_ColReg31", i, oRecordSet01.Fields.Item("BaseIDX").Value.ToString().Trim()); //기준계획번호

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                PH_PY204_Add_MatrixRow(oMat01.VisualRowCount, false);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PH_PY204_AddData()
        {
            bool functionReturnValue = false;

            short i;
            string sQry;

            string IDX; //IDX
            string MSTCOD; //대상자사번
            string MSTNAM; //대상자성명
            string StdYear; //기준년도
            string TeamCode; //부서코드
            string TeamName; //부서명
            string RspCode; //담당코드
            string RspName; //담당명
            string ClsCode; //반코드
            string ClsName; //반명
            string JIGNAM; //직급
            string Class; //교육구분
            string ClassName; //교육구분명
            string EduName; //교육명
            string EduOrg; //교육기관
            string EduFrDt; //교육시작일
            string EduToDt; //교육종료일
            string EduHour; //교육시간
            decimal EduAmount; //교육비
            string EduQuarter; //분기
            string EduHarf; //반기
            string EduMonth; //월
            string Reference; //관련근거
            string Comment; //비고
            string CLTCOD; //사업장
            string DocStatus; //상태
            string PlnCls; //계획구분
            string Certi; //이수증(수료증)
            string Report; //보고서
            string EduPlnCnt; //교육계획건수
            string BaseIDX; //기준교육계획번호(교육계획 수정시 등록)

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, false);

            try
            {
                oMat01.FlushToDataSource();
                //마지막 빈행 제외를 위해 2를 뺌
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (oDS_PH_PY204B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        IDX = (oDS_PH_PY204B.GetValue("U_ColReg02", i).ToString().Trim() == "" ? "-1" : oDS_PH_PY204B.GetValue("U_ColReg02", i).ToString().Trim());
                        MSTCOD = oDS_PH_PY204B.GetValue("U_ColReg04", i).ToString().Trim(); //대상자사번
                        MSTNAM = oDS_PH_PY204B.GetValue("U_ColReg05", i).ToString().Trim(); //대상자성명
                        StdYear = oDS_PH_PY204B.GetValue("U_ColReg03", i).ToString().Trim(); //기준년도
                        TeamCode = oDS_PH_PY204B.GetValue("U_ColReg06", i).ToString().Trim(); //부서코드
                        TeamName = oDS_PH_PY204B.GetValue("U_ColReg07", i).ToString().Trim(); //부서명
                        RspCode = oDS_PH_PY204B.GetValue("U_ColReg08", i).ToString().Trim(); //담당코드
                        RspName = oDS_PH_PY204B.GetValue("U_ColReg09", i).ToString().Trim(); //담당명
                        ClsCode = oDS_PH_PY204B.GetValue("U_ColReg10", i).ToString().Trim(); //반코드
                        ClsName = oDS_PH_PY204B.GetValue("U_ColReg11", i).ToString().Trim(); //반명
                        JIGNAM = oDS_PH_PY204B.GetValue("U_ColReg12", i).ToString().Trim(); //직급
                        Class = oDS_PH_PY204B.GetValue("U_ColReg13", i).ToString().Trim(); //교육구분
                        ClassName = oDS_PH_PY204B.GetValue("U_ColReg14", i).ToString().Trim(); //교육구분명
                        EduName = oDS_PH_PY204B.GetValue("U_ColReg15", i).ToString().Trim(); //교육명
                        EduOrg = oDS_PH_PY204B.GetValue("U_ColReg16", i).ToString().Trim(); //교육기관
                        EduFrDt = oDS_PH_PY204B.GetValue("U_ColDt01", i).ToString().Trim(); //교육시작일
                        EduToDt = oDS_PH_PY204B.GetValue("U_ColDt02", i).ToString().Trim(); //교육종료일
                        EduHour = oDS_PH_PY204B.GetValue("U_ColQty01", i).ToString().Trim(); //교육시간
                        EduAmount = Convert.ToDecimal(oDS_PH_PY204B.GetValue("U_ColSum03", i).ToString()); //교육비
                        EduQuarter = oDS_PH_PY204B.GetValue("U_ColSum01", i).ToString().Trim(); //분기
                        EduHarf = oDS_PH_PY204B.GetValue("U_ColReg18", i).ToString().Trim(); //반기
                        EduMonth = oDS_PH_PY204B.GetValue("U_ColSum02", i).ToString().Trim(); //월
                        Reference = oDS_PH_PY204B.GetValue("U_ColReg24", i).ToString().Trim(); //관련근거
                        Comment = oDS_PH_PY204B.GetValue("U_ColReg25", i).ToString().Trim(); //비고
                        CLTCOD = oDS_PH_PY204B.GetValue("U_ColReg26", i).ToString().Trim(); //사업장
                        DocStatus = oDS_PH_PY204B.GetValue("U_ColReg27", i).ToString().Trim(); //상태
                        PlnCls = oDS_PH_PY204B.GetValue("U_ColReg28", i).ToString().Trim(); //계획구분
                        Certi = oDS_PH_PY204B.GetValue("U_ColReg29", i).ToString().Trim(); //이수증
                        Report = oDS_PH_PY204B.GetValue("U_ColReg30", i).ToString().Trim(); //보고서
                        EduPlnCnt = oDS_PH_PY204B.GetValue("U_ColSum04", i).ToString().Trim(); //교육계획건수
                        BaseIDX = oDS_PH_PY204B.GetValue("U_ColReg31", i).ToString().Trim(); //기준계획번호

                        sQry = "EXEC [PH_PY204_02] '";
                        sQry += IDX + "','"; //IDX
                        sQry += MSTCOD + "','"; //대상자사번
                        sQry += MSTNAM + "','"; //대상자성명
                        sQry += StdYear + "','"; //기준년도
                        sQry += TeamCode + "','"; //부서코드
                        sQry += TeamName + "','"; //부서명
                        sQry += RspCode + "','"; //담당코드
                        sQry += RspName + "','"; //담당명
                        sQry += ClsCode + "','"; //반코드
                        sQry += ClsName + "','"; //반명
                        sQry += JIGNAM + "','"; //직급
                        sQry += Class + "','"; //교육구분
                        sQry += ClassName + "','"; //교육구분명
                        sQry += EduName + "','"; //교육명
                        sQry += EduOrg + "','"; //교육기관
                        sQry += EduFrDt + "','"; //교육시작일
                        sQry += EduToDt + "',"; //교육종료일
                        sQry += EduHour + ","; //교육시간
                        sQry += EduAmount + ","; //교육비
                        sQry += EduQuarter + ",'"; //분기
                        sQry += EduHarf + "',"; //반기
                        sQry += EduMonth + ",'"; //월
                        sQry += Reference + "','"; //관련근거
                        sQry += Comment + "','"; //비고
                        sQry += CLTCOD + "','"; //사업장
                        sQry += DocStatus + "','"; //상태
                        sQry += PlnCls + "','"; //계획구분
                        sQry += Certi + "','"; //이수증(수료증)
                        sQry += Report + "','"; //보고서
                        sQry += EduPlnCnt + "','"; //교육계획건수
                        sQry += BaseIDX + "'"; //기준계획번호

                        RecordSet01.DoQuery(sQry);

                        ProgBar01.Value += 1;
                        ProgBar01.Text = ProgBar01.Value + "/" + (oMat01.VisualRowCount - 1) + "건 저장중...";
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_AddData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 기본정보 삭제
        /// </summary>
        private void PH_PY204_DeleteData()
        {
            short loopCount = 0;
            short ErrNum = 0;
            string sQry;
            string IDX;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("삭제 중...", 0, false);

            try
            {
                oMat01.FlushToDataSource();

                //마직막 빈행 제외를 위해 2를 뺌
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
                {
                    if (oDS_PH_PY204B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        IDX = oDS_PH_PY204B.GetValue("U_ColReg02", loopCount).ToString().Trim(); //Index

                        //삭제 가능 여부 체크
                        sQry = "EXEC PH_PY204_51 ";
                        sQry += IDX;

                        oRecordSet01.DoQuery(sQry);

                        sQry = "";

                        if (oRecordSet01.Fields.Item("ReturnValue").Value == "False")
                        {
                            ErrNum = 2;
                            throw new Exception();
                        }

                        sQry = "EXEC PH_PY204_04 ";
                        sQry += IDX;

                        oRecordSet01.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("삭제대상이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(loopCount + 1 + "행은 실적이 등록되어 삭제가 불가능합니다. " + loopCount + 1 + "행 이후 삭제는 모두 취소됩니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_DeleteData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// 메트릭스 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PH_PY204_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            int i = 0;
            short ErrNum = 0;

            try
            {
                oMat01.FlushToDataSource();

                for (i = 0; i <= oMat01.VisualRowCount - 2; i++) //마지막 빈행 제외를 위해 2를 뺌
                {
                    if (oDS_PH_PY204B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {

                        if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg04", i).ToString().Trim())) //사번
                        {
                            if (oDS_PH_PY204B.GetValue("U_ColReg13", i).ToString().Trim() != "12" && oDS_PH_PY204B.GetValue("U_ColReg13", i).ToString().Trim() != "9") //특별교육 Or 직무교육은 사번 체크 제외
                            {
                                ErrNum = 1;
                                throw new Exception();
                            }
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg05", i).ToString().Trim())) //성명
                        {
                            ErrNum = 2;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg03", i).ToString().Trim())) //기준년도
                        {
                            ErrNum = 3;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg13", i).ToString().Trim())) //구분
                        {
                            ErrNum = 4;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg15", i).ToString().Trim())) //교육명
                        {
                            ErrNum = 5;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColDt01", i).ToString().Trim())) //교육시작일
                        {
                            ErrNum = 6;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColDt02", i).ToString().Trim())) //교육종료일
                        {
                            ErrNum = 7;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg06", i).ToString().Trim())) //부서
                        {
                            ErrNum = 8;
                            throw new Exception();
                        }
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 사번이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 성명이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 기준년도가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 구분이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육명이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육시작일이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육종료일이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 8)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 부서코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY204_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int loopCount;
            string sQry;

            string CLTCOD;
            string TeamCode; //대상자부서
            string TeamName;
            string RspCode; //대상자담당
            string RspName;
            string ClsCode; //대상자반
            string ClsName;
            string FullName; //성명
            string EduMonth; //교육월
            string EduQuarter = string.Empty; //교육분기
            string EduHarf = string.Empty; //교육반기
            string JIGNAM; //직급

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "Mat01":

                        oMat01.FlushToDataSource();

                        if (oCol == "MSTCOD")
                        {
                            oDS_PH_PY204B.SetValue("U_ColReg04", oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim()); //사번
                            oDS_PH_PY204B.SetValue("U_ColReg03", oRow - 1, oForm.Items.Item("StdYear").Specific.Value.ToString().Trim()); //기준년도

                            //대상자의 인사마스터에서 소속 조회
                            sQry = "  SELECT  T0.U_TeamCode AS [TeamCode], "; //부서코드
                            sQry += "         T1.U_CodeNm AS [TeamName], "; //부서명
                            sQry += "         T0.U_RspCode AS [RspCode], "; //담당코드
                            sQry += "         T2.U_CodeNm AS [RspName], "; //담당명
                            sQry += "         T0.U_ClsCode AS [ClsCode], "; //반코드
                            sQry += "         T3.U_CodeNm AS [ClsName], "; //반명
                            sQry += "         T0.U_FullName AS [FullName], "; //성명
                            sQry += "         T4.U_CodeNm AS [JIGNAM],"; //직급
                            sQry += "         T0.U_CLTCOD AS [CLTCOD]"; //소속사업장
                            sQry += " FROM    [@PH_PY001A] AS T0 ";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T1";
                            sQry += "             ON T0.U_TeamCode = T1.U_Code";
                            sQry += "             AND T1.Code = '1'";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T2";
                            sQry += "             ON T0.U_RspCode = T2.U_Code";
                            sQry += "             AND T2.Code = '2'";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T3";
                            sQry += "             ON T0.U_ClsCode = T3.U_Code";
                            sQry += "             AND T3.Code = '9'";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T4";
                            sQry += "             ON T0.U_JIGCOD = T4.U_Code";
                            sQry += "             AND T4.Code = 'P129'";
                            sQry += " WHERE   T0.Code = '" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);

                            TeamCode = oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim();
                            TeamName = oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim();
                            RspCode = oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim();
                            RspName = oRecordSet01.Fields.Item("RspName").Value.ToString().Trim();
                            ClsCode = oRecordSet01.Fields.Item("ClsCode").Value.ToString().Trim();
                            ClsName = oRecordSet01.Fields.Item("ClsName").Value.ToString().Trim();
                            FullName = (oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() == "9999999" ? "전사원" : oRecordSet01.Fields.Item("FullName").Value.ToString().Trim());
                            JIGNAM = oRecordSet01.Fields.Item("JIGNAM").Value.ToString().Trim();
                            CLTCOD = oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim();

                            oDS_PH_PY204B.SetValue("U_ColReg06", oRow - 1, TeamCode); //부서코드
                            oDS_PH_PY204B.SetValue("U_ColReg07", oRow - 1, TeamName); //부서명
                            oDS_PH_PY204B.SetValue("U_ColReg08", oRow - 1, RspCode); //담당코드
                            oDS_PH_PY204B.SetValue("U_ColReg09", oRow - 1, RspName); //담당명
                            oDS_PH_PY204B.SetValue("U_ColReg10", oRow - 1, ClsCode); //반코드
                            oDS_PH_PY204B.SetValue("U_ColReg11", oRow - 1, ClsName); //반명
                            oDS_PH_PY204B.SetValue("U_ColReg05", oRow - 1, FullName); //성명
                            oDS_PH_PY204B.SetValue("U_ColReg12", oRow - 1, JIGNAM); //직급
                            oDS_PH_PY204B.SetValue("U_ColReg26", oRow - 1, CLTCOD); //사업장

                            //행 추가
                            if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PH_PY204B.GetValue("U_ColReg04", oRow - 1).ToString().Trim()))
                            {
                                PH_PY204_Add_MatrixRow(oRow, false);
                            }

                            oMat01.Columns.Item("JIGNAM").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        }
                        else if (oCol == "TeamCode")
                        {
                            oDS_PH_PY204B.SetValue("U_ColReg07", oRow - 1, (oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString.Trim() == "9999" ? "전부서" : dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'", " AND Code = '1'"))); //부서
                        }
                        else if (oCol == "RspCode")
                        {
                            oDS_PH_PY204B.SetValue("U_ColReg09", oRow - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'", " AND Code = '2'")); //담당
                        }
                        else if (oCol == "ClsCode")
                        {
                            oDS_PH_PY204B.SetValue("U_ColReg11", oRow - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'", " AND Code = '9'")); //반
                        }
                        else if (oCol == "Class")
                        {
                            oDS_PH_PY204B.SetValue("U_ColReg14", oRow - 1, dataHelpClass.Get_ReData("name", "edType", "OHED", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'", "")); //교육구분
                        }
                        else if (oCol == "EduFrDt") //교육시작일
                        {
                            EduMonth = oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Substring(4, 2);

                            if (Convert.ToInt16(EduMonth) >= 1 && Convert.ToInt16(EduMonth) <= 3)
                            {
                                EduQuarter = "1";
                                EduHarf = "상";
                            }
                            else if (Convert.ToInt16(EduMonth) >= 4 && Convert.ToInt16(EduMonth) <= 6)
                            {
                                EduQuarter = "2";
                                EduHarf = "상";
                            }
                            else if (Convert.ToInt16(EduMonth) >= 7 && Convert.ToInt16(EduMonth) <= 9)
                            {
                                EduQuarter = "3";
                                EduHarf = "하";
                            }
                            else if (Convert.ToInt16(EduMonth) >= 10 & Convert.ToInt16(EduMonth) <= 12)
                            {
                                EduQuarter = "4";
                                EduHarf = "하";
                            }

                            oDS_PH_PY204B.SetValue("U_ColSum01", oRow - 1, Convert.ToString(Convert.ToInt16(EduQuarter))); //분기
                            oDS_PH_PY204B.SetValue("U_ColReg18", oRow - 1, EduHarf); //반기
                            oDS_PH_PY204B.SetValue("U_ColSum02", oRow - 1, Convert.ToString(Convert.ToInt16(EduMonth))); //월
                        }

                        oMat01.LoadFromDataSource();
                        //강제 포커스 이동_S
                        oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //강제 포커스 이동_E
                        oMat01.AutoResizeColumns();
                        break;

                    case "CLTCOD":

                        CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                        {
                            for (loopCount = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                            {
                                oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //부서콤보세팅
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "  SELECT      U_Code AS [Code],";
                        sQry += "             U_CodeNm As [Name]";
                        sQry += " FROM        [@PS_HR200L]";
                        sQry += " WHERE       Code = '1'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += "             AND U_Char2 = '" + CLTCOD + "'";
                        sQry += " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "MSTCOD":

                        oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", ""); //성명
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Matrix 체크박스 전체 선택
        /// </summary>
        private void PH_PY204_CheckAll()
        {
            string CheckType;
            short loopCount;

            CheckType = "Y";

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY204B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY204B.Offset = loopCount;
                    if (CheckType == "N")
                    {
                        oDS_PH_PY204B.SetValue("U_ColReg01", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY204B.SetValue("U_ColReg01", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_CheckAll_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 엑셀자료 업로드
        /// </summary>
        private void PH_PY204_UploadExcel()
        {
            int rowCount;
            int loopCount;
            string sFile;
            
            bool sucessFlag = false;
            short columnCount = 29; //엑셀 컬럼수

            FileListBoxForm fileListBoxForm = new FileListBoxForm();

            //*.xls|*.xls|*.xlsx|*.xlsx| 'Dialog창의 "선택콤보박스내용|미리보기창의 파일 Filter" 한 쌍임
            sFile = fileListBoxForm.OpenDialog(fileListBoxForm, "*.xls|*.xls|*.xlsx|*.xlsx|", "파일선택", "C:\\");

            if (string.IsNullOrEmpty(sFile))
            {
                return;
            }

            //엑셀 Object 연결
            //암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
            Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);
            Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
            Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
            Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
            Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);

            oForm.Freeze(true);

            oMat01.Clear();
            oMat01.FlushToDataSource();
            oMat01.LoadFromDataSource();

            try
            {
                for (rowCount = 2; rowCount <= xlRow.Count; rowCount++)
                {
                    if (rowCount - 2 != 0)
                    {
                        oDS_PH_PY204B.InsertRecord(rowCount - 2);
                    }

                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[rowCount, loopCount];
                    }

                    oDS_PH_PY204B.Offset = rowCount - 2;
                    oDS_PH_PY204B.SetValue("U_LineNum", rowCount - 2, Convert.ToString(rowCount - 1));
                    oDS_PH_PY204B.SetValue("U_ColReg01", rowCount - 2, "Y");
                    oDS_PH_PY204B.SetValue("U_ColReg04", rowCount - 2, Convert.ToString(r[1].Value)); //사번
                    oDS_PH_PY204B.SetValue("U_ColReg05", rowCount - 2, Convert.ToString(r[2].Value)); //성명
                    oDS_PH_PY204B.SetValue("U_ColReg03", rowCount - 2, Convert.ToString(r[3].Value)); //기준년도
                    oDS_PH_PY204B.SetValue("U_ColReg06", rowCount - 2, Convert.ToString(r[4].Value)); //부서코드
                    oDS_PH_PY204B.SetValue("U_ColReg07", rowCount - 2, Convert.ToString(r[5].Value)); //부서명
                    oDS_PH_PY204B.SetValue("U_ColReg08", rowCount - 2, Convert.ToString(r[6].Value)); //담당코드
                    oDS_PH_PY204B.SetValue("U_ColReg09", rowCount - 2, Convert.ToString(r[7].Value)); //담당명
                    oDS_PH_PY204B.SetValue("U_ColReg10", rowCount - 2, Convert.ToString(r[8].Value)); //반코드
                    oDS_PH_PY204B.SetValue("U_ColReg11", rowCount - 2, Convert.ToString(r[9].Value)); //반명
                    oDS_PH_PY204B.SetValue("U_ColReg12", rowCount - 2, Convert.ToString(r[10].Value)); //직급
                    oDS_PH_PY204B.SetValue("U_ColReg13", rowCount - 2, Convert.ToString(r[11].Value)); //교육구분
                    oDS_PH_PY204B.SetValue("U_ColReg14", rowCount - 2, Convert.ToString(r[12].Value)); //교육구분명
                    oDS_PH_PY204B.SetValue("U_ColReg15", rowCount - 2, Convert.ToString(r[13].Value)); //교육명
                    oDS_PH_PY204B.SetValue("U_ColReg16", rowCount - 2, Convert.ToString(r[14].Value)); //교육기관
                    oDS_PH_PY204B.SetValue("U_ColDt01", rowCount - 2, Convert.ToDateTime(Convert.ToString(r[15].Value)).ToString("yyyyMMdd")); //교육시작일
                    oDS_PH_PY204B.SetValue("U_ColDt02", rowCount - 2, Convert.ToDateTime(Convert.ToString(r[16].Value)).ToString("yyyyMMdd")); //교육종료일
                    oDS_PH_PY204B.SetValue("U_ColQty01", rowCount - 2, Convert.ToString(r[17].Value)); //교육시간
                    oDS_PH_PY204B.SetValue("U_ColSum03", rowCount - 2, Convert.ToString(r[18].Value)); //교육비
                    oDS_PH_PY204B.SetValue("U_ColSum01", rowCount - 2, Convert.ToString(r[19].Value)); //분기
                    oDS_PH_PY204B.SetValue("U_ColReg18", rowCount - 2, Convert.ToString(r[20].Value)); //반기
                    oDS_PH_PY204B.SetValue("U_ColSum02", rowCount - 2, Convert.ToString(r[21].Value)); //월
                    oDS_PH_PY204B.SetValue("U_ColReg24", rowCount - 2, Convert.ToString(r[22].Value)); //관련근거
                    oDS_PH_PY204B.SetValue("U_ColReg25", rowCount - 2, Convert.ToString(r[23].Value)); //비고
                    oDS_PH_PY204B.SetValue("U_ColReg26", rowCount - 2, Convert.ToString(r[24].Value)); //사업장
                    oDS_PH_PY204B.SetValue("U_ColReg27", rowCount - 2, Convert.ToString(r[25].Value)); //상태
                    oDS_PH_PY204B.SetValue("U_ColReg28", rowCount - 2, Convert.ToString(r[26].Value)); //계획구분
                    oDS_PH_PY204B.SetValue("U_ColReg29", rowCount - 2, Convert.ToString(r[27].Value)); //이수증(수료증)
                    oDS_PH_PY204B.SetValue("U_ColReg30", rowCount - 2, Convert.ToString(r[28].Value)); //보고서
                    oDS_PH_PY204B.SetValue("U_ColSum04", rowCount - 2, Convert.ToString(r[29].Value)); //건수

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
                    }
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();

                PH_PY204_Add_MatrixRow(oMat01.RowCount, false);
                sucessFlag = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox("[PH_PY204_UploadExcel_Error]" + (char)13 + ex.Message);
            }
            finally
            {
                //액셀개체 닫음
                xlapp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                oForm.Freeze(false);
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

                ////case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                ////    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                ////    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

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
                ////    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                ////    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                ////    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    //    break;

                //    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    //    break;
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
                    if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY204_MatrixSpaceLineDel() == false) //매트릭스 필수자료 체크
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PH_PY204_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY204_LoadCaption();
                            PH_PY204_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                        PH_PY204_LoadCaption();
                        PH_PY204_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
                        {
                            PH_PY204_DeleteData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                            PH_PY204_LoadCaption();
                            PH_PY204_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSel") //전체선택
                    {
                        PH_PY204_CheckAll();
                    }
                    else if (pVal.ItemUID == "BtnExcel") //엑셀업로드
                    {
                        PH_PY204_UploadExcel();
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                    }
                    else if (pVal.ItemUID == "BtnPrint2")
                    {
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY204")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", ""); //Header-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MSTCOD"); //Matrix-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "TeamCode"); //Matrix-부서
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "RspCode"); //Matrix-담당
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ClsCode"); //Matrix-반
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "Class"); //Matrix-교육구분
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
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else
                    {
                        PH_PY204_FlushToItemValue(pVal.ItemUID, 0, "");
                    }
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "MSTCOD" || pVal.ColUID == "TeamCode" || pVal.ColUID == "RspCode" || pVal.ColUID == "ClsCode" || pVal.ColUID == "Class" || pVal.ColUID == "EduFrDt")
                            {
                                PH_PY204_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                        }
                        else
                        {
                            PH_PY204_FlushToItemValue(pVal.ItemUID, 0, "");
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
                    PH_PY204_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY204B);
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
                    PH_PY204_FormResize();
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
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY204_FormReset();
                            PH_PY204_LoadCaption();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기
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
                        case "7169": //엑셀 내보내기
                            break;
                    }
                }
            }
            catch (Exception ex)
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

