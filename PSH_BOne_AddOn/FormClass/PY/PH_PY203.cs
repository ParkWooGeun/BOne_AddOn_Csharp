using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 교육실적등록
    /// </summary>
    internal class PH_PY203 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Matrix oMat01;

        private SAPbouiCOM.DBDataSource oDS_PH_PY203B; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private int oLast_Mode;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY203.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY203_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY203");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                ////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

                oForm.Freeze(true);
                PH_PY203_CreateItems();
                PH_PY203_ComboBox_Setting();
                PH_PY203_CF_ChooseFromList();
                PH_PY203_EnableMenus();
                PH_PY203_SetDocument(oFormDocEntry01);
                PH_PY203_FormResize();
                PH_PY203_Add_MatrixRow(0, true);
                PH_PY203_LoadCaption();
                PH_PY203_FormReset();
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
        private void PH_PY203_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY203B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

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

                //교육명
                oForm.DataSources.UserDataSources.Add("EduName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("EduName").Specific.DataBind.SetBound(true, "", "EduName");

                //교육기관
                oForm.DataSources.UserDataSources.Add("EduOrg", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("EduOrg").Specific.DataBind.SetBound(true, "", "EduOrg");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY203_ComboBox_Setting()
        {
            string sQry = string.Empty;

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
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm AS [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = '4'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";

                oForm.Items.Item("StdMt").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("StdMt").Specific, sQry, "", false, false);
                oForm.Items.Item("StdMt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //매트릭스-분류
                sQry = "        SELECT      edType AS [Code],";
                sQry = sQry + "             name As [Name]";
                sQry = sQry + " FROM        [OHED]";
                sQry = sQry + " ORDER BY    edType";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("EduType"), sQry, "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(true);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PH_PY203_CF_ChooseFromList()
        {
            try
            {
                oForm.Freeze(true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_CF_ChooseFromList_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY203_EnableMenus()
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY203_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY203_FormItemEnabled();
                    ////Call PH_PY203_AddMatrixRow(0, True) '//UDO방식일때
                }
                else
                {
                    //        oForm.Mode = fm_FIND_MODE
                    //        Call PH_PY203_FormItemEnabled
                    //        oForm.Items("DocEntry").Specific.Value = oFormDocEntry01
                    //        oForm.Items("1").Click ct_Regular
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY203_FormItemEnabled()
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY203_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_FormResize_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY203_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PH_PY203B.InsertRecord((oRow));
                }

                oMat01.AddRow();
                oDS_PH_PY203B.Offset = oRow;
                oDS_PH_PY203B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PH_PY203B.SetValue("U_ColReg01", oRow, "Y");

                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_Add_MatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>        
        private void PH_PY203_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    //        oForm.Items("BtnDelete").Enabled = False
                    //    ElseIf oForm.Mode = fm_OK_MODE Then
                    //        oForm.Items("BtnAdd").Specific.Caption = "확인"
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    //        oForm.Items("BtnDelete").Enabled = True
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_LoadCaption_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PH_PY203_FormReset()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);

                string User_BPLID = null;
                User_BPLID = dataHelpClass.User_BPLID();

                //헤더 초기화
                oForm.DataSources.UserDataSources.Item("CLTCOD").Value = User_BPLID; //사업장
                oForm.DataSources.UserDataSources.Item("TeamCode").Value = "%"; //부서
                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = ""; //사번
                oForm.DataSources.UserDataSources.Item("MSTNAM").Value = ""; //성명
                oForm.DataSources.UserDataSources.Item("Status").Value = "1"; //재직여부
                oForm.DataSources.UserDataSources.Item("StdYear").Value = DateTime.Now.ToString("yyyy"); //기준년도
                oForm.DataSources.UserDataSources.Item("EduName").Value = ""; //교육명
                oForm.DataSources.UserDataSources.Item("EduOrg").Value = ""; //교육기관

                //라인 초기화
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                PH_PY203_Add_MatrixRow(0, true);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY203_MTX01()
        {
            int i = 0;
            string sQry = string.Empty;
            short ErrNum = 0;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string MSTCOD = string.Empty; //사번
            string StdYear = string.Empty; //기준년도
            string StdMt = string.Empty; //기준월
            string Status = string.Empty; //재직여부
            string EduName = string.Empty; //교육명
            string EduOrg = string.Empty; //교육기관

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
            TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //부서
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //사번
            StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim(); //기준년도
            StdMt = oForm.Items.Item("StdMt").Specific.Value.ToString().Trim(); //기준월
            Status = oForm.Items.Item("Status").Specific.Value.ToString().Trim(); //재직여부
            EduName = oForm.Items.Item("EduName").Specific.Value.ToString().Trim(); //교육명
            EduOrg = oForm.Items.Item("EduOrg").Specific.Value.ToString().Trim(); //교육기관

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 100, false);

            try
            {
                oForm.Freeze(true);

                sQry = "      EXEC [PH_PY203_01] '";
                sQry = sQry + CLTCOD + "','"; //사업장
                sQry = sQry + TeamCode + "','"; //부서
                sQry = sQry + MSTCOD + "','"; //사번
                sQry = sQry + StdYear + "','"; //기준년도
                sQry = sQry + StdMt + "','"; //기준월
                sQry = sQry + Status + "','"; //재직여부
                sQry = sQry + EduName + "','"; //교육명
                sQry = sQry + EduOrg + "'"; //교육기관

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY203B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY203_Add_MatrixRow(0, true);
                    PH_PY203_LoadCaption();

                    ErrNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY203B.Size)
                    {
                        oDS_PH_PY203B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY203B.Offset = i;

                    oDS_PH_PY203B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY203B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim()); //선택
                    oDS_PH_PY203B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //사번
                    oDS_PH_PY203B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim()); //성명
                    oDS_PH_PY203B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("EduType").Value.ToString().Trim()); //분류
                    oDS_PH_PY203B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("EduName").Value.ToString().Trim()); //교육명
                    oDS_PH_PY203B.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet01.Fields.Item("FrDt").Value.ToString()).ToString("yyyyMMdd")); //기간시작
                    oDS_PH_PY203B.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet01.Fields.Item("ToDt").Value.ToString()).ToString("yyyyMMdd")); //기간종료
                    oDS_PH_PY203B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("EduOrg").Value.ToString().Trim()); //교육기관
                    oDS_PH_PY203B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("EduLoc").Value.ToString().Trim()); //교육장소
                    oDS_PH_PY203B.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("EduHour").Value.ToString().Trim()); //교육시간
                    oDS_PH_PY203B.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("EduAmt").Value.ToString().Trim()); //교육비
                    oDS_PH_PY203B.SetValue("U_ColSum02", i, oRecordSet01.Fields.Item("EduAmt2").Value.ToString().Trim()); //출장비
                    oDS_PH_PY203B.SetValue("U_ColRgL01", i, oRecordSet01.Fields.Item("Comment").Value.ToString().Trim()); //비고
                    oDS_PH_PY203B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("PlnIDX").Value.ToString().Trim()); //교육계획번호
                    oDS_PH_PY203B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("LineID").Value.ToString().Trim()); //LineID

                    //Call oDS_PH_PY203B.setValue("U_ColSum06", i, Trim(oRecordSet01.Fields("TotalExp").Value)) '합계

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                PH_PY203_Add_MatrixRow(oMat01.VisualRowCount, false);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                ProgBar01.Stop();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private bool PH_PY203_AddData()
        {
            bool functionReturnValue = false;
            
            short i = 0;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("저장 중...", 100, false);

            string CLTCOD = string.Empty; //사업장코드
            string MSTCOD = string.Empty; //사번
            string MSTNAM = string.Empty; //성명
            string EduType = string.Empty; //분류
            string EduName = string.Empty; //교육명
            string FrDt = string.Empty; //기간시작
            string ToDt = string.Empty; //기간종료
            string EduOrg = string.Empty; //교육기관
            string EduLoc = string.Empty; //교육장소
            double EduHour = 0; //교육시간
            double EduAmt = 0; //교육비
            double EduAmt2 = 0; //출장비
            string Comment = string.Empty; //비고
            string PlnIDX = string.Empty; //교육계획번호
            string LineId = string.Empty; //라인ID

            try
            {
                oMat01.FlushToDataSource();
                //마직막 빈행 제외를 위해 2를 뺌
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (oDS_PH_PY203B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        CLTCOD = oForm.DataSources.UserDataSources.Item("CLTCOD").Value; //사업장코드
                        MSTCOD = oDS_PH_PY203B.GetValue("U_ColReg02", i).ToString().Trim(); //사번
                        MSTNAM = oDS_PH_PY203B.GetValue("U_ColReg03", i).ToString().Trim(); //성명
                        EduType = oDS_PH_PY203B.GetValue("U_ColReg04", i).ToString().Trim(); //분류
                        EduName = oDS_PH_PY203B.GetValue("U_ColReg05", i).ToString().Trim(); //교육명
                        FrDt = oDS_PH_PY203B.GetValue("U_ColDt01", i).ToString().Trim(); //기간시작
                        ToDt = oDS_PH_PY203B.GetValue("U_ColDt02", i).ToString().Trim(); //기간종료
                        EduOrg = oDS_PH_PY203B.GetValue("U_ColReg08", i).ToString().Trim(); //교육기관
                        EduLoc = oDS_PH_PY203B.GetValue("U_ColReg09", i).ToString().Trim(); //교육장소
                        EduHour = Convert.ToDouble(oDS_PH_PY203B.GetValue("U_ColQty01", i).ToString().Trim()); //교육시간
                        EduAmt = Convert.ToDouble(oDS_PH_PY203B.GetValue("U_ColSum01", i).ToString().Trim()); //교육비
                        EduAmt2 = Convert.ToDouble(oDS_PH_PY203B.GetValue("U_ColSum02", i).ToString().Trim()); //출장비
                        Comment = oDS_PH_PY203B.GetValue("U_ColRgL01", i).ToString().Trim(); //비고
                        PlnIDX = oDS_PH_PY203B.GetValue("U_ColReg10", i).ToString().Trim(); //교육계획번호
                        LineId = string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg12", i).ToString().Trim()) ? "-1" : oDS_PH_PY203B.GetValue("U_ColReg12", i).ToString().Trim(); //라인ID

                        sQry = "            EXEC [PH_PY203_02] ";
                        sQry = sQry + "'" + CLTCOD + "',"; //사업장
                        sQry = sQry + "'" + MSTCOD + "',"; //사번
                        sQry = sQry + "'" + MSTNAM + "',"; //성명
                        sQry = sQry + "'" + EduType + "',"; //분류
                        sQry = sQry + "'" + EduName + "',"; //교육명
                        sQry = sQry + "'" + FrDt + "',"; //기간시작
                        sQry = sQry + "'" + ToDt + "',"; //기간종료
                        sQry = sQry + "'" + EduOrg + "',"; //교육기관
                        sQry = sQry + "'" + EduLoc + "',"; //교육장소
                        sQry = sQry + "'" + EduHour + "',"; //교육시간
                        sQry = sQry + "'" + EduAmt + "',"; //교육비
                        sQry = sQry + "'" + EduAmt2 + "',"; //출장비
                        sQry = sQry + "'" + Comment + "',"; //비고
                        sQry = sQry + "'" + PlnIDX + "',"; //교육계획번호
                        sQry = sQry + "'" + LineId + "'"; //라인ID

                        RecordSet01.DoQuery(sQry);
                    }
                }

                ProgBar01.Value = 100;
                ProgBar01.Stop();

                PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_AddData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY203_DeleteData()
        {
            short i = 0;
            string sQry = string.Empty;
            short ErrNum = 0;

            string MSTCOD = string.Empty;
            string LineId = string.Empty;
            string PlnIDX = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("삭제 중...", 100, false);

            try
            {
                oMat01.FlushToDataSource();
                //마직막 빈행 제외를 위해 2를 뺌
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (oDS_PH_PY203B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        MSTCOD = oDS_PH_PY203B.GetValue("U_ColReg02", i).ToString().Trim(); //사번
                        LineId = string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg12", i).ToString().Trim()) ? "-1" : oDS_PH_PY203B.GetValue("U_ColReg12", i).ToString().Trim(); //라인ID
                        PlnIDX = oDS_PH_PY203B.GetValue("U_ColReg10", i).ToString().Trim(); //교육계획번호

                        sQry = "            EXEC PH_PY203_04 ";
                        sQry = sQry + "'" + MSTCOD + "',"; //사번
                        sQry = sQry + "'" + LineId + "',"; //라인ID
                        sQry = sQry + "'" + PlnIDX + "'"; //교육계획번호

                        oRecordSet01.DoQuery(sQry);
                    }
                }

                ProgBar01.Value = 100;
                ProgBar01.Stop();

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("삭제대상이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_DeleteData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private bool PH_PY203_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            int i = 0;
            short ErrNum = 0;
            
            try
            {
                oMat01.FlushToDataSource();
                
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++) //마지막 빈행 제외를 위해 2를 뺌
                {
                    if (oDS_PH_PY203B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg02", i).ToString().Trim())) //사번
                        {
                            ErrNum = 1;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg03", i).ToString().Trim())) //성명
                        {
                            ErrNum = 2;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg04", i).ToString().Trim())) //분류
                        {
                            ErrNum = 3;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg05", i).ToString().Trim())) //교육명
                        {
                            ErrNum = 4;
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColDt01", i).ToString().Trim())) //기간시작
                        {
                            ErrNum = 5;
                            throw new Exception();

                        }
                        else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColDt02", i).ToString().Trim())) //기간종료
                        {
                            ErrNum = 6;
                            throw new Exception();
                        }
                        //else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg08", i).ToString().Trim())) //교육기관
                        //{
                        //    ErrNum = 7;
                        //    throw new Exception();

                        //}
                        //else if (string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg09", i).ToString().Trim())) //교육장소
                        //{
                        //    ErrNum = 8;
                        //    throw new Exception();
                        //}
                    }
                }

                functionReturnValue = true;
            }
            catch(Exception ex)
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 분류가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육명이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 기간시작이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 기간종료가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                //else if (ErrNum == 7)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육기관이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
                //else if (ErrNum == 8)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 교육장소 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                functionReturnValue = false;
            }


            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY203_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int loopCount = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty; //사번
            string MSTNAM = string.Empty; //성명
            string EduType = string.Empty; //교육구분
            string EduName = string.Empty; //교육명
            string EduFrDt = string.Empty; //교육시작일
            string EduToDt = string.Empty; //교육종료일
            string EduOrg = string.Empty; //교육기관
            string EduLoc = string.Empty; //교육장소
            double EduHour = 0; //교육시간
            double EduAmount = 0; //교육비
            double EduAmount2 = 0; //출장비(0)
            string Comment = string.Empty; //비고('')

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
                            oDS_PH_PY203B.SetValue("U_ColReg02", oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value);
                            oDS_PH_PY203B.SetValue("U_ColReg03", oRow - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value + "'", ""));
                            if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg02", oRow - 1).ToString().Trim()))
                            {
                                PH_PY203_Add_MatrixRow(oRow, false);
                            }
                        }
                        else if (oCol == "PlnIDX")
                        {
                            //교육계획에서 자료 조회
                            sQry = "        SELECT  T0.MSTCOD AS [MSTCOD], "; //사번
                            sQry = sQry + "         T0.MSTNAM AS [MSTNAM], "; //성명
                            sQry = sQry + "         T0.Class AS [EduType], "; //교육구분
                            sQry = sQry + "         T0.EduName AS [EduName], "; //교육명
                            sQry = sQry + "         T0.EduFrDt AS [EduFrDt], "; //교육시작일
                            sQry = sQry + "         T0.EduToDt AS [EduToDt], "; //교육종료일
                            sQry = sQry + "         T0.EduOrg AS [EduOrg], "; //교육기관
                            sQry = sQry + "         '' AS [EduLoc], "; //교육장소
                            sQry = sQry + "         T0.EduHour AS [EduHour], "; //교육시간
                            sQry = sQry + "         T0.EduAmount AS [EduAmount], "; //교육비
                            sQry = sQry + "         0 AS [EduAmount2], "; //출장비
                            sQry = sQry + "         '' AS [Comment] "; //비고
                            sQry = sQry + " FROM    Z_PH_PY204 AS T0 ";
                            sQry = sQry + " WHERE   T0.IDX = '" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value + "'";

                            oRecordSet01.DoQuery(sQry);

                            MSTCOD = oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim(); //사번
                            MSTNAM = oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim(); //성명
                            EduType = oRecordSet01.Fields.Item("EduType").Value.ToString().Trim(); //교육구분
                            EduName = oRecordSet01.Fields.Item("EduName").Value.ToString().Trim(); //교육명
                            EduFrDt = oRecordSet01.Fields.Item("EduFrDt").Value.ToString("yyyyMMdd"); //교육시작일
                            EduToDt = oRecordSet01.Fields.Item("EduToDt").Value.ToString("yyyyMMdd"); //교육종료일
                            EduOrg = oRecordSet01.Fields.Item("EduOrg").Value.ToString().Trim(); //교육기관
                            EduLoc = oRecordSet01.Fields.Item("EduLoc").Value.ToString().Trim(); //교육장소
                            EduHour = Convert.ToDouble(oRecordSet01.Fields.Item("EduHour").Value.ToString().Trim()); //교육시간
                            EduAmount = Convert.ToDouble(oRecordSet01.Fields.Item("EduAmount").Value.ToString().Trim()); //교육비
                            EduAmount2 = Convert.ToDouble(oRecordSet01.Fields.Item("EduAmount2").Value.ToString().Trim()); //출장비
                            Comment = oRecordSet01.Fields.Item("Comment").Value.ToString().Trim(); //비고

                            oDS_PH_PY203B.SetValue("U_ColReg02", oRow - 1, MSTCOD); //사번
                            oDS_PH_PY203B.SetValue("U_ColReg03", oRow - 1, MSTNAM); //성명
                            oDS_PH_PY203B.SetValue("U_ColReg04", oRow - 1, EduType); //교육구분
                            oDS_PH_PY203B.SetValue("U_ColReg05", oRow - 1, EduName); //교육명
                            oDS_PH_PY203B.SetValue("U_ColDt01", oRow - 1, EduFrDt); //교육시작일
                            oDS_PH_PY203B.SetValue("U_ColDt02", oRow - 1, EduToDt); //교육종료일
                            oDS_PH_PY203B.SetValue("U_ColReg08", oRow - 1, EduOrg); //교육기관
                            oDS_PH_PY203B.SetValue("U_ColReg09", oRow - 1, EduLoc); //교육장소
                            oDS_PH_PY203B.SetValue("U_ColQty01", oRow - 1, Convert.ToString(EduHour)); //교육시간
                            oDS_PH_PY203B.SetValue("U_ColSum01", oRow - 1, Convert.ToString(EduAmount)); //교육비
                            oDS_PH_PY203B.SetValue("U_ColSum02", oRow - 1, Convert.ToString(EduAmount2)); //출장비
                            oDS_PH_PY203B.SetValue("U_ColRgL01", oRow - 1, Comment); //비고

                            if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PH_PY203B.GetValue("U_ColReg10", oRow - 1).ToString().Trim()))
                            {
                                PH_PY203_Add_MatrixRow(oRow, false);
                            }
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
                        sQry = "        SELECT      U_Code AS [Code],";
                        sQry = sQry + "             U_CodeNm As [Name]";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '1'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + "             AND U_Char2 = '" + CLTCOD + "'";
                        sQry = sQry + " ORDER BY    U_Seq";
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY203_CheckAll()
        {
            string CheckType = string.Empty;
            short loopCount = 0;

            CheckType = "Y";

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY203B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY203B.Offset = loopCount;
                    if (CheckType == "N")
                    {
                        oDS_PH_PY203B.SetValue("U_ColReg01", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY203B.SetValue("U_ColReg01", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY203_CheckAll_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 엑셀자료 업로드
        /// </summary>
        private void PH_PY203_UploadExcel()
        {
            int rowCount = 0;
            int loopCount = 0;
            string sFile = string.Empty;

            bool sucessFlag = false;
            short columnCount = 13; //엑셀 컬럼수

            FileListBoxForm fileListBoxForm = new FileListBoxForm();

            //*.xls|*.xls|*.xlsx|*.xlsx| 'Dialog창의 "선택콤보박스내용|미리보기창의 파일 Filter" 한 쌍임
            sFile = fileListBoxForm.OpenDialog(fileListBoxForm, "*.xls|*.xls|*.xlsx|*.xlsx|", "파일선택", "C:\\");

            if (string.IsNullOrEmpty(sFile))
            {
                //PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                        oDS_PH_PY203B.InsertRecord(rowCount - 2);
                    }

                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[rowCount, loopCount];
                    }

                    oDS_PH_PY203B.Offset = rowCount - 2;
                    oDS_PH_PY203B.SetValue("U_LineNum", rowCount - 2, Convert.ToString(rowCount - 1));
                    oDS_PH_PY203B.SetValue("U_ColReg01", rowCount - 2, "Y");
                    oDS_PH_PY203B.SetValue("U_ColReg02", rowCount - 2, Convert.ToString(r[1].Value));
                    oDS_PH_PY203B.SetValue("U_ColReg03", rowCount - 2, Convert.ToString(r[2].Value));
                    oDS_PH_PY203B.SetValue("U_ColReg04", rowCount - 2, Convert.ToString(r[3].Value));
                    oDS_PH_PY203B.SetValue("U_ColReg05", rowCount - 2, Convert.ToString(r[4].Value));
                    oDS_PH_PY203B.SetValue("U_ColDt01", rowCount - 2, Convert.ToDateTime(Convert.ToString(r[5].Value)).ToString("yyyyMMdd"));
                    oDS_PH_PY203B.SetValue("U_ColDt02", rowCount - 2, Convert.ToDateTime(Convert.ToString(r[6].Value)).ToString("yyyyMMdd"));
                    oDS_PH_PY203B.SetValue("U_ColReg08", rowCount - 2, Convert.ToString(r[7].Value));
                    oDS_PH_PY203B.SetValue("U_ColReg09", rowCount - 2, Convert.ToString(r[8].Value));
                    oDS_PH_PY203B.SetValue("U_ColQty01", rowCount - 2, Convert.ToString(r[9].Value));
                    oDS_PH_PY203B.SetValue("U_ColSum01", rowCount - 2, Convert.ToString(r[10].Value));
                    oDS_PH_PY203B.SetValue("U_ColSum02", rowCount - 2, Convert.ToString(r[11].Value));
                    oDS_PH_PY203B.SetValue("U_ColRgL01", rowCount - 2, Convert.ToString(r[12].Value));
                    oDS_PH_PY203B.SetValue("U_ColReg10", rowCount - 2, Convert.ToString(r[13].Value));

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
                    }
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();

                PH_PY203_Add_MatrixRow(oMat01.RowCount, false);
                sucessFlag = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox("[PH_PY203_UploadExcel_Error]" + (char)13 + ex.Message);
                sucessFlag = false;
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

                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

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
                    if (pVal.ItemUID == "PH_PY203")
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

                    if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {

                            //                If PH_PY203_HeaderSpaceLineDel() = False Then
                            //                    BubbleEvent = False
                            //                    Exit Sub
                            //                End If

                            //                If PH_PY203_DataCheck() = False Then
                            //                    BubbleEvent = False
                            //                    Exit Sub
                            //                End If

                            //매트릭스 필수자료 체크
                            if (PH_PY203_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PH_PY203_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //                Call PH_PY203_FormReset
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY203_LoadCaption();
                            PH_PY203_MTX01();

                            oLast_Mode = (int)oForm.Mode;

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                            //if (PH_PY203_HeaderSpaceLineDel() == false)
                            //{
                            //    BubbleEvent = false;
                            //    return;
                            //}

                            ////                If PH_PY203_DataCheck() = False Then
                            ////                    BubbleEvent = False
                            ////                    Exit Sub
                            ////                End If

                            //if (PH_PY203_UpdateData() == false)
                            //{
                            //    BubbleEvent = false;
                            //    return;
                            //}

                            //PH_PY203_FormReset();
                            //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            //PH_PY203_LoadCaption();
                            //PH_PY203_MTX01();

                            ////                oForm.Items("GCode").Click ct_Regular
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {

                        //PH_PY203_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                        PH_PY203_LoadCaption();
                        PH_PY203_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
                        {
                            PH_PY203_DeleteData();
                            //PH_PY203_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                            PH_PY203_LoadCaption();
                            PH_PY203_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSel") //전체선택
                    {
                        PH_PY203_CheckAll();
                    }
                    else if (pVal.ItemUID == "BtnExcel") //엑셀업로드
                    {
                        PH_PY203_UploadExcel();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY203")
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PlnIDX"); //Matrix-교육계획번호
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
                        PH_PY203_FlushToItemValue(pVal.ItemUID, 0, "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            if (pVal.ColUID == "MSTCOD" || pVal.ColUID == "PlnIDX")
                            {
                                PH_PY203_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                        }
                        else
                        {
                            PH_PY203_FlushToItemValue(pVal.ItemUID, 0, "");
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY203_FormItemEnabled();
                    ////Call PH_PY203_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY203B);
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
                    PH_PY203_FormResize();
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
                    //        If (pval.ItemUID = "ItemCode") Then
                    //            Dim oDataTable01 As SAPbouiCOM.DataTable
                    //            Set oDataTable01 = pval.SelectedObjects
                    //            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
                    //            Set oDataTable01 = Nothing
                    //        End If
                    //        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
                    //            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY203A", "U_CardCode,U_CardName")
                    //        End If
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        	oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        	PH_PY203_FormReset();
                        	PH_PY203_LoadCaption();
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
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
