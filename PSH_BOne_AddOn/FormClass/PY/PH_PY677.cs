using System;

using SAPbouiCOM;
using SAPbobsCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 일일근태이상자조회
    /// </summary>
    internal class PH_PY677 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Matrix oMat01;

        private SAPbouiCOM.DBDataSource oDS_PH_PY677B; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY677.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY677_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY677");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                ////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

                oForm.Freeze(true);
                PH_PY677_CreateItems();
                PH_PY677_CF_ChooseFromList();
                PH_PY677_EnableMenus();
                PH_PY677_SetDocument(oFromDocEntry01);
                PH_PY677_FormResize();

                //PH_PY677_Add_MatrixRow(0, true);
                //PH_PY677_LoadCaption();
                //PH_PY677_FormItemEnabled();

                oForm.Items.Item("MSTCOD").Click(); //사번 포커스
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
        private void PH_PY677_CreateItems()
        {
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY677B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                // 기간
                oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");
                oForm.Items.Item("FrDate").Specific.String = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");
                oForm.Items.Item("ToDate").Specific.String = DateTime.Now.ToString("yyyyMMdd");

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                // 반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                // 사원번호(코드)
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 사원번호(성명)
                oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

                // 근무형태(코드)
                oForm.DataSources.UserDataSources.Add("ShiftDatCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ShiftDatCd").Specific.DataBind.SetBound(true, "", "ShiftDatCd");

                // 근무형태(명)
                oForm.DataSources.UserDataSources.Add("ShiftDatNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ShiftDatNm").Specific.DataBind.SetBound(true, "", "ShiftDatNm");

                // 근무조(코드)
                oForm.DataSources.UserDataSources.Add("GNMUJOCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("GNMUJOCd").Specific.DataBind.SetBound(true, "", "GNMUJOCd");

                // 근무조(명)
                oForm.DataSources.UserDataSources.Add("GNMUJONm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("GNMUJONm").Specific.DataBind.SetBound(true, "", "GNMUJONm");

                // 기찰이상구분
                oForm.DataSources.UserDataSources.Add("Class", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Class").Specific.DataBind.SetBound(true, "", "Class");
                oForm.Items.Item("Class").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Class").Specific.ValidValues.Add("Y", "근태이상");
                oForm.Items.Item("Class").Specific.ValidValues.Add("N", "정상");
                oForm.Items.Item("Class").DisplayDesc = true;
                oForm.Items.Item("Class").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 근태기찰정상확인
                oForm.DataSources.UserDataSources.Add("Confirm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Confirm").Specific.DataBind.SetBound(true, "", "Confirm");
                oForm.Items.Item("Confirm").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Confirm").Specific.ValidValues.Add("N", "미확인[N]");
                oForm.Items.Item("Confirm").Specific.ValidValues.Add("Y", "확인[Y]");
                oForm.Items.Item("Confirm").DisplayDesc = true;
                oForm.Items.Item("Confirm").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 근태구분
                oForm.DataSources.UserDataSources.Add("WorkType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("WorkType").Specific.DataBind.SetBound(true, "", "WorkType");
                
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P221'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("WorkType").Specific, "N");
                oForm.Items.Item("WorkType").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("WorkType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //oForm.Items.Item("WorkType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 매트릭스 기본값 SETTING

                // 근무형태
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P154'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat01.Columns.Item("ShiftDat").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.Columns.Item("ShiftDat").DisplayDesc = true;

                // 근무조
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P155'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat01.Columns.Item("GNMUJO").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.Columns.Item("GNMUJO").DisplayDesc = true;

                // 요일구분
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P202'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat01.Columns.Item("DayType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.Columns.Item("DayType").DisplayDesc = true;

                // 근태구분
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P221'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat01.Columns.Item("P_WorkType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.Columns.Item("P_WorkType").DisplayDesc = true;

                // 기찰정상확인
                oMat01.Columns.Item("P_Confirm").ValidValues.Add("N", "미확인[N]");
                oMat01.Columns.Item("P_Confirm").ValidValues.Add("Y", "확인[Y]");
                oMat01.Columns.Item("P_Confirm").DisplayDesc = true;

                // 근태이상분류
                sQry = "        SELECT    U_Code AS [Code],";
                sQry = sQry + "           U_CodeNm As [Name]";
                sQry = sQry + " FROM      [@PS_HR200L]";
                sQry = sQry + " WHERE     Code = 'P237'";
                sQry = sQry + "           AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY  U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat01.Columns.Item("WkAbCls").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.Columns.Item("WkAbCls").DisplayDesc = true;

                // 교대인정
                oMat01.Columns.Item("RotateYN").ValidValues.Add("N", "N");
                oMat01.Columns.Item("RotateYN").ValidValues.Add("Y", "Y");
                oMat01.Columns.Item("RotateYN").DisplayDesc = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY677_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PH_PY677_CF_ChooseFromList()
        {
            try
            {
                oForm.Freeze(true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_CF_ChooseFromList_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY677_EnableMenus()
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY011_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY677_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY677_FormItemEnabled();
                    ////Call PH_PY677_AddMatrixRow(0, True) '//UDO방식일때
                }
                else
                {
                    //        oForm.Mode = fm_FIND_MODE
                    //        Call PH_PY677_FormItemEnabled
                    //        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
                    //        oForm.Items("1").Click ct_Regular
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY677_FormItemEnabled()
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY677_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_FormResize_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경(사용 안함, 호환성을 위해 남겨둠)
        /// </summary>
        private void PH_PY677_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    oForm.Items.Item("BtnDelete").Enabled = false;
                    //    ElseIf oForm.Mode = fm_OK_MODE Then
                    //        oForm.Items("BtnAdd").Specific.Caption = "확인"
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    oForm.Items.Item("BtnDelete").Enabled = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_LoadCaption_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY677_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PH_PY677B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY677B.Offset = oRow;
                oDS_PH_PY677B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_Add_MatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY677_MTX01()
        {
            short i = 0;
            string sQry = string.Empty;
            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 100, false); ;

            string CLTCOD = string.Empty;            // 사업장
            string FrDate = string.Empty;            // 시작일자
            string ToDate = string.Empty;            // 종료일자
            string TeamCode = string.Empty;          // 부서
            string RspCode = string.Empty;           // 담당
            string ClsCode = string.Empty;           // 반
            string ShiftDat = string.Empty;          // 근무형태
            string GNMUJO = string.Empty;            // 근무조
            string MSTCOD = string.Empty;            // 사원번호
            string Class_Renamed = string.Empty;     // 기찰이상구분(2014.04.10 송명규 추가)
            string Confirm = string.Empty;           // 근태기찰정상확인(2013.03.29 송명규 추가)
            string WorkType = string.Empty;          // 근태구분(2014.05.13 송명규 추가)


            try
            {
                oForm.Freeze(true);

                CLTCOD =   oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                FrDate =   oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim();
                ToDate =   oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.VALUE.ToString().Trim();
                RspCode =  oForm.Items.Item("RspCode").Specific.VALUE.ToString().Trim();
                ClsCode =  oForm.Items.Item("ClsCode").Specific.VALUE.ToString().Trim();
                ShiftDat = oForm.Items.Item("ShiftDatCd").Specific.VALUE.ToString().Trim();
                GNMUJO =   oForm.Items.Item("GNMUJOCd").Specific.VALUE.ToString().Trim();
                MSTCOD =   oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                Class_Renamed = oForm.Items.Item("Class").Specific.VALUE.ToString().Trim();
                Confirm =  oForm.Items.Item("Confirm").Specific.VALUE.ToString().Trim();
                WorkType = oForm.Items.Item("WorkType").Specific.VALUE.ToString().Trim();

                sQry = "            EXEC [PH_PY677_01] ";
                sQry = sQry + "'" + CLTCOD + "',";                // 사업장
                sQry = sQry + "'" + FrDate + "',";                // 시작일자
                sQry = sQry + "'" + ToDate + "',";                // 종료일자
                sQry = sQry + "'" + TeamCode + "',";              // 부서
                sQry = sQry + "'" + RspCode + "',";               // 담당
                sQry = sQry + "'" + ClsCode + "',";               // 반
                sQry = sQry + "'" + ShiftDat + "',";              // 근무형태
                sQry = sQry + "'" + GNMUJO + "',";                // 근무조
                sQry = sQry + "'" + MSTCOD + "',";                // 사원번호
                sQry = sQry + "'" + Class_Renamed + "',";         // 기찰이상구분
                sQry = sQry + "'" + Confirm + "',";               // 근태기찰정상확인
                sQry = sQry + "'" + WorkType + "'";               // 근태구분

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY677B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY677B.Size)
                    {
                        oDS_PH_PY677B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY677B.Offset = i;

                    oDS_PH_PY677B.SetValue("U_LineNum",  i, Convert.ToString(i + 1));
                    oDS_PH_PY677B.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("Chk").Value);                          //선택
                    oDS_PH_PY677B.SetValue("U_ColDt05",  i, Convert.ToDateTime(oRecordSet01.Fields.Item("PosDate").Value.ToString().Trim()).ToString("yyyyMMdd"));   //일자
                    oDS_PH_PY677B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("TeamName").Value);                     //부서
                    oDS_PH_PY677B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("RspName").Value);                      //담당
                    oDS_PH_PY677B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ClsName").Value);                      //반
                    oDS_PH_PY677B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("MSTCOD").Value);                       //사번
                    oDS_PH_PY677B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("MSTNAM").Value);                       //성명
                    oDS_PH_PY677B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("ShiftDat").Value);                     //근무형태
                    oDS_PH_PY677B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("GNMUJO").Value);                       //근무조
                    oDS_PH_PY677B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("DayWeek").Value);                      //요일
                    oDS_PH_PY677B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("DayType").Value);                      //요일구분
                    oDS_PH_PY677B.SetValue("U_ColDt01",  i, Convert.ToDateTime(oRecordSet01.Fields.Item("P_GetDt").Value.ToString().Trim()).ToString("yyyyMMdd"));   //출근일자(계획)
                    oDS_PH_PY677B.SetValue("U_ColTm01",  i, oRecordSet01.Fields.Item("P_GetTime").Value);                    //출근시각(계획)
                    oDS_PH_PY677B.SetValue("U_ColDt02",  i, Convert.ToDateTime(oRecordSet01.Fields.Item("P_OffDt").Value.ToString().Trim()).ToString("yyyyMMdd"));   //퇴근일자(계획)
                    oDS_PH_PY677B.SetValue("U_ColTm02",  i, oRecordSet01.Fields.Item("P_OffTime").Value);                    //퇴근시각(계획)
                    oDS_PH_PY677B.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("P_Base").Value);                       //기본(계획)
                    oDS_PH_PY677B.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("P_Extend").Value);                     //연장(계획)
                    oDS_PH_PY677B.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("P_MidNt").Value);                      //심야
                    oDS_PH_PY677B.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("P_Special").Value);                    //특근(계획)
                    oDS_PH_PY677B.SetValue("U_ColQty05", i, oRecordSet01.Fields.Item("P_SpExtend").Value);                   //특연(계획)
                    oDS_PH_PY677B.SetValue("U_ColQty06", i, oRecordSet01.Fields.Item("P_EarlyTo").Value);                    //조출
                    oDS_PH_PY677B.SetValue("U_ColQty07", i, oRecordSet01.Fields.Item("P_SEarlyTo").Value);                   //휴조
                    oDS_PH_PY677B.SetValue("U_ColQty08", i, oRecordSet01.Fields.Item("P_EducTran").Value);                   //교육훈련
                    oDS_PH_PY677B.SetValue("U_ColQty09", i, oRecordSet01.Fields.Item("P_LateTo").Value);                     //지각
                    oDS_PH_PY677B.SetValue("U_ColQty10", i, oRecordSet01.Fields.Item("P_EarlyOff").Value);                   //조퇴
                    oDS_PH_PY677B.SetValue("U_ColReg21", i, oRecordSet01.Fields.Item("P_WorkType").Value);                   //근태구분
                    oDS_PH_PY677B.SetValue("U_ColReg22", i, oRecordSet01.Fields.Item("P_Comment").Value);                    //비고
                    oDS_PH_PY677B.SetValue("U_ColQty11", i, oRecordSet01.Fields.Item("P_GoOut").Value);                      //외출
                    oDS_PH_PY677B.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("P_Confirm").Value);                    //기찰정상확인
                    oDS_PH_PY677B.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("R_GetDt").Value);                      //출근일자(기찰)
                    oDS_PH_PY677B.SetValue("U_ColTm03",  i, oRecordSet01.Fields.Item("R_GetTime").Value);                    //출근시각(기찰)
                    oDS_PH_PY677B.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("R_OffDt").Value);                      //퇴근일자(기찰)
                    oDS_PH_PY677B.SetValue("U_ColTm04",  i, oRecordSet01.Fields.Item("R_OffTime").Value);                    //퇴근시각(기찰)

                    //        Call oDS_PH_PY677B.setValue("U_ColQty05", i, Trim(oRecordSet01.Fields("R_Base").VALUE)) '기본(기찰)
                    //        Call oDS_PH_PY677B.setValue("U_ColQty06", i, Trim(oRecordSet01.Fields("R_Extend").VALUE)) '연장(기찰)
                    //        Call oDS_PH_PY677B.setValue("U_ColQty07", i, Trim(oRecordSet01.Fields("R_Special").VALUE)) '특근(기찰)
                    //        Call oDS_PH_PY677B.setValue("U_ColQty08", i, Trim(oRecordSet01.Fields("R_SpExtend").VALUE)) '특연(기찰)
                    //        Call oDS_PH_PY677B.setValue("U_ColQty09", i, Trim(oRecordSet01.Fields("R_TotTime").VALUE)) '총근무시간(기찰)

                    oDS_PH_PY677B.SetValue("U_ColQty12", i, oRecordSet01.Fields.Item("Rotation").Value);                    //교대일수
                    oDS_PH_PY677B.SetValue("U_ColReg24", i, oRecordSet01.Fields.Item("R_YN").Value);                        //기찰완료여부
                    oDS_PH_PY677B.SetValue("U_ColReg25", i, oRecordSet01.Fields.Item("WkAbCls").Value);                     //근태이상분류
                    oDS_PH_PY677B.SetValue("U_ColReg26", i, oRecordSet01.Fields.Item("WkAbCmt").Value);                     //근태이상사유
                    oDS_PH_PY677B.SetValue("U_ColReg27", i, oRecordSet01.Fields.Item("ActText").Value);                     //근무내용
                    oDS_PH_PY677B.SetValue("U_ColReg28", i, oRecordSet01.Fields.Item("RotateYN").Value);                    //교대인정

                    oRecordSet01.MoveNext();

                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                ProgBar01.Stop();

            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// 기본정보 수정
        /// </summary>
        /// <returns></returns>
        private bool PH_PY677_UpdateData()
        {
            bool functionReturnValue = false;

            int i = 0;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = null;
            RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            string CLTCOD = string.Empty;            // 사업장
            string PosDate = string.Empty;           // 일자
            string MSTCOD = string.Empty;            // 사번
            string JIGTYP = string.Empty;            // 지급타입
            string PAYTYP = string.Empty;            // 급여타입
            string JIGCOD = string.Empty;            // 직급코드
            string ShiftDat = string.Empty;          // 근무형태
            string GNMUJO = string.Empty;            // 근무조
            string P_WorkType = string.Empty;        // 근태구분
            string P_Confirm = string.Empty;         // 기찰정상확인
            string P_GetTime = string.Empty;         // 출근시각
            string P_OffDt = string.Empty;           // 퇴근일자
            string P_OffTime = string.Empty;         // 퇴근시각
            double P_Base = 0;                       // 기본
            double P_Extend = 0;                     // 연장
            double P_Special = 0;                    // 특근
            double P_SpExtend = 0;                   // 특연
            double P_Midnight = 0;                   // 심야
            double P_EarlyTo = 0;                    // 조출
            double P_SEarlyTo = 0;                   // 휴조
            double P_EducTran = 0;                   // 교육훈련
            double P_LateTo = 0;                     // 지각
            double P_EarlyOff = 0;                   // 조퇴
            double P_GoOut = 0;                      // 외출
            string P_Comment = string.Empty;         // 비고
            string DangerCd = string.Empty;          // 비고
            string WkAbCls = string.Empty;           // 근태이상분류
            string WkAbCmt = string.Empty;           // 근태이상사유
            string ActText = string.Empty;           // 근무내용
            string RotateYN = string.Empty;          // 교대인정

            try
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("수정 중...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (oDS_PH_PY677B.GetValue("U_ColReg20", i).ToString().Trim() == "Y")
                    {
                        CLTCOD     = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();                     // 사업장
                        PosDate    = oDS_PH_PY677B.GetValue("U_ColDt05", i).Trim();                        // 일자
                        MSTCOD     = oDS_PH_PY677B.GetValue("U_ColReg05", i).Trim();                       // 사번
                        ShiftDat   = oDS_PH_PY677B.GetValue("U_ColReg07", i).Trim();                       // 근무형태
                        GNMUJO     = oDS_PH_PY677B.GetValue("U_ColReg08", i).Trim();                       // 근무조
                        P_WorkType = oDS_PH_PY677B.GetValue("U_ColReg21", i).Trim();                       // 근태구분
                        P_Confirm  = oDS_PH_PY677B.GetValue("U_ColReg16", i).Trim();                       // 기찰정상확인
                        P_GetTime  = oDS_PH_PY677B.GetValue("U_ColTm01", i).Trim();                        // 출근시각
                        P_OffDt    = oDS_PH_PY677B.GetValue("U_ColDt02", i).Trim();                        // 퇴근일자
                        P_OffTime  = oDS_PH_PY677B.GetValue("U_ColTm02", i).Trim();                        // 퇴근시각
                        P_Base     = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty01", i));            // 기본
                        P_Extend   = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty02", i));            // 연장
                        P_Special  = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty04", i));            // 특근
                        P_SpExtend = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty05", i));            // 특연
                        P_Midnight = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty03", i));            // 심야
                        P_EarlyTo  = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty06", i));            // 조출
                        P_SEarlyTo = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty07", i));            // 휴조
                        P_EducTran = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty08", i));            // 교육훈련
                        P_LateTo   = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty09", i));            // 지각
                        P_EarlyOff = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty10", i));            // 조퇴
                        P_GoOut    = Convert.ToDouble(oDS_PH_PY677B.GetValue("U_ColQty11", i));            // 외출
                        P_Comment = oDS_PH_PY677B.GetValue("U_ColReg22", i).Trim();                        // 비고
                        WkAbCls   = oDS_PH_PY677B.GetValue("U_ColReg25", i).Trim();                        // 근태이상분류
                        WkAbCmt   = oDS_PH_PY677B.GetValue("U_ColReg26", i).Trim();                        // 근태이상사유
                        ActText   = oDS_PH_PY677B.GetValue("U_ColReg27", i).Trim();                        // 근무내용
                        RotateYN  = oDS_PH_PY677B.GetValue("U_ColReg28", i).Trim();                        // 교대인정
                        //DangerCd = dataHelpClass.Get_ReData("DangerCD", "PosDate", "ZPH_PY008", "'" + Convert.ToDateTime(oDS_PH_PY677B.GetValue("U_ColDt05", i)) + "'", " and mstcod ='" + MSTCOD + "'");
                        DangerCd = dataHelpClass.Get_ReData("DangerCD", "PosDate", "ZPH_PY008", "'" + oDS_PH_PY677B.GetValue("U_ColDt05", i) + "'", " and mstcod ='" + MSTCOD + "'");

                        sQry = "Select   U_JIGTYP";
                        sQry = sQry + ", U_PAYTYP";
                        sQry = sQry + ", U_JIGCOD ";
                        sQry = sQry + "from [@PH_PY001A] ";
                        sQry = sQry + "where u_status <> '5' and code ='" + MSTCOD + "' ";

                        RecordSet01.DoQuery(sQry);

                        JIGTYP = RecordSet01.Fields.Item(0).Value.Trim();
                        PAYTYP = RecordSet01.Fields.Item(1).Value.Trim();
                        JIGCOD = RecordSet01.Fields.Item(2).Value.Trim();

                        // 무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                        if (P_WorkType == "A01" | P_WorkType == "A02" | P_WorkType.Substring(0, 1) == "F" | P_WorkType == "D11")
                        {
                            DangerCd = "";
                        }
                        else
                        {
                            // 전문직, 계약직 이며 연봉제가 아니고 위해코드가 없으면 위해코드를 기타로..
                            if ((JIGTYP == "04" | JIGTYP == "05") & PAYTYP != "1" & JIGCOD != "73" & string.IsNullOrEmpty(DangerCd))
                            {
                                // 창원사업장만 적용(2013.09.30 송명규 추가)
                                if (CLTCOD == "1")
                                {
                                    DangerCd = "31";
                                }
                            }
                        }

                        sQry = "            EXEC [PH_PY677_02] ";
                        sQry = sQry + "'" + CLTCOD + "',";                        // 사업장
                        sQry = sQry + "'" + PosDate + "',";                       // 일자
                        sQry = sQry + "'" + MSTCOD + "',";                        // 사번
                        sQry = sQry + "'" + ShiftDat + "',";                      // 근무형태
                        sQry = sQry + "'" + GNMUJO + "',";                        // 근무조
                        sQry = sQry + "'" + P_WorkType + "',";                    // 근태구분
                        sQry = sQry + "'" + P_Confirm + "',";                     // 기찰정상확인
                        sQry = sQry + "'" + P_GetTime + "',";                     // 출근시각
                        sQry = sQry + "'" + P_OffDt + "',";                       // 퇴근일자
                        sQry = sQry + "'" + P_OffTime + "',";                     // 퇴근시각
                        sQry = sQry + "'" + P_Base + "',";                        // 기본
                        sQry = sQry + "'" + P_Extend + "',";                      // 연장
                        sQry = sQry + "'" + P_Special + "',";                     // 특근
                        sQry = sQry + "'" + P_SpExtend + "',";                    // 특연
                        sQry = sQry + "'" + P_Midnight + "',";                    // 심야
                        sQry = sQry + "'" + P_EarlyTo + "',";                     // 조출
                        sQry = sQry + "'" + P_SEarlyTo + "',";                    // 휴조
                        sQry = sQry + "'" + P_EducTran + "',";                    // 교육훈련
                        sQry = sQry + "'" + P_LateTo + "',";                      // 지각
                        sQry = sQry + "'" + P_EarlyOff + "',";                    // 조퇴
                        sQry = sQry + "'" + P_GoOut + "',";                       // 외출
                        sQry = sQry + "'" + P_Comment + "',";                     // 비고
                        sQry = sQry + "'" + WkAbCls + "',";                       // 근태이상분류
                        sQry = sQry + "'" + WkAbCmt + "',";                       // 근태이상사유
                        sQry = sQry + "'" + ActText + "',";                       // 근무내용
                        sQry = sQry + "'" + RotateYN + "',";                      // 교대인정 황영수(2019.01.31)
                        sQry = sQry + "'" + DangerCd + "'";                       // 위해코드

                        RecordSet01.DoQuery(sQry);
                    }
                }
                PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_UpdateData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PH_PY677_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short ErrNum = 0;

            try
            {
                //if (oForm.Items.Item("DestNo1").Specific.Value.Trim() == "") //출장번호1
                //{
                //    ErrNum = 1;
                //    throw new Exception();
                //}
                //else if (oForm.Items.Item("DestNo2").Specific.Value.Trim() == "") //출장번호2
                //{
                //    ErrNum = 2;
                //    throw new Exception();
                //}

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("출장번호1은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DestNo1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("출장번호2는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DestNo2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_HeaderSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                functionReturnValue = false;
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 메트릭스 필수 사항 check
        /// 구현은 되어 있지만 사용하지 않음
        /// </summary>
        /// <returns></returns>
        private bool PH_PY677_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;

            int i = 0;
            short ErrNum = 0;

            try
            {
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY677_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD, TeamCode, RspCode = string.Empty;
            string PreWorkType = string.Empty;
            string WorkType = string.Empty;
            string ymd = string.Empty;
            string MSTCOD = string.Empty;
            string YY = string.Empty;
            Double JanQty = 0;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            SAPbobsCOM.Recordset oRecordSet01 = null;
            oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //oForm.Freeze(true);

                switch (oUID)
                {
                    case "Mat01":

                        PreWorkType = oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1).ToString().Trim();

                        oMat01.FlushToDataSource();

                        if (oCol == "P_GetTime")
                        {
                            PH_PY677_Time_ReSet(oRow);
                            oMat01.LoadFromDataSource();
                            oMat01.Columns.Item(oCol).Cells.Item(oRow).Click();

                        }
                        else if (oCol == "P_OffTime")
                        {

                            if (oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1) != "0000")
                            {
                                PH_PY677_Time_Calc_Main(oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1), oRow);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(oCol).Cells.Item(oRow).Click();
                            }

                        }
                        else if (oCol == "P_WorkType")
                        {

                            WorkType = oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1).ToString().Trim();

                            switch (WorkType)
                            {
                                case "A01":
                                case "E02":
                                case "E03":
                                case "F01":
                                case "F02":
                                case "F03":
                                case "F04":
                                case "F05":
                                    {
                                        // 무단결근, 유계결근, 무급휴일, 휴업, 병가(휴직), 신병휴직, 정직(유결), 가사휴직, 공상휴직(F01) 추가(2017.12.07)
                                        oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)); // 퇴근일자(P_OffDt)
                                        oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00"); // 출근시각(P_GetTime)
                                        oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00"); // 퇴근시각(P_OffTime)
                                        PH_PY677_Time_ReSet(oRow);
                                        oMat01.LoadFromDataSource();
                                        break;
                                    }

                                case "C02":
                                case "D04":
                                case "D05":
                                case "D06":
                                case "D07":
                                case "H05":
                                    {
                                        // 훈련, 경조휴가, 하기휴가, 특별휴가, 분만휴가, 조합활동
                                        oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1)); // 퇴근일자(P_OffDt)
                                        oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");   // 출근시각(P_GetTime)
                                        oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");   // 퇴근시각(P_OffTime)
                                        oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(0)); // 교대일수
                                        PH_PY677_Time_ReSet(oRow);
                                        oMat01.LoadFromDataSource();
                                        break;
                                    }
                                case "D02":
                                case "D09":
                                    {
                                        // 연차/반차 휴가
                                        // 연차/반차 휴가 잔여일수 확인
                                        if (WorkType == "D02")
                                        {
                                            JanQty = 1;
                                        }
                                        else if (WorkType == "D09")
                                        {
                                            JanQty = 0.5;
                                        }

                                        ymd = oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1).ToString().Trim();
                                        CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                                        MSTCOD = oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1).ToString().Trim();

                                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회중...", oRecordSet01.RecordCount, false);

                                        sQry = "      EXEC [PH_PY775_01] '";
                                        sQry = sQry + CLTCOD + "','";
                                        sQry = sQry + ymd.Substring(0,4) + "','";
                                        sQry = sQry + MSTCOD + "'";

                                        oRecordSet01.DoQuery(sQry);

                                        if (oRecordSet01.Fields.Item("jandd").Value < JanQty)
                                        {
                                            ErrNum = 1;
                                            oDS_PH_PY677B.SetValue("U_ColReg21", oRow - 1, "A00");
                                            oMat01.LoadFromDataSource();
                                            throw new Exception();
                                        }
                                        else
                                        {
                                            oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));  // 퇴근일자(P_OffDt)
                                            oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");                                        // 출근시각(P_GetTime)
                                            oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");                                        // 퇴근시각(P_OffTime)
                                            oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(1));                           // 교대일수

                                            PH_PY677_Time_ReSet(oRow);
                                            oMat01.LoadFromDataSource();
                                        }

                                        ProgressBar01.Value = 100;
                                        ProgressBar01.Stop();
                                        break;
                                    }
                                case "D08":
                                case "D10":
                                    {
                                        //근속보전휴가, 근속보전반차(기계사업부)
                                        //근속보전휴가 잔량 확인
                                        ymd = oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1).ToString().Trim();
                                        CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                                        MSTCOD = oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1).ToString().Trim();

                                        if (Convert.ToDateTime(dataHelpClass.ConvertDateType(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1), "-")) >= Convert.ToDateTime(ymd.Substring(0, 4) + "-07-01") )
                                        {
                                            YY = ymd.Substring(0, 4);
                                        }
                                        else
                                        {
                                            YY = Convert.ToString(Convert.ToInt16(ymd.Substring(0, 4)) - 1);
                                        }

                                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회중...", oRecordSet01.RecordCount, false);

                                        sQry = "        SELECT     Bqty = ISNULL(( ";
                                        sQry = sQry + "                               SELECT     ISNULL(c.Qty,0) ";
                                        sQry = sQry + "                               FROM       [ZPH_GUNSOKCHA] c";
                                        sQry = sQry + "                               WHERE      c.Year = '" + YY + "'";
                                        sQry = sQry + "                                          AND (";
                                        sQry = sQry + "                                                   CASE ";
                                        sQry = sQry + "                                                       WHEN ISDATE(a.U_GrpDat) = 0 THEN a.U_StartDat ";
                                        sQry = sQry + "                                                       ELSE A.U_GrpDat ";
                                        sQry = sQry + "                                                   END";
                                        sQry = sQry + "                                              ) BETWEEN c.DocDateFr AND c.DocDateTo";
                                        sQry = sQry + "                           ),0), ";                                        //발생수량
                                        sQry = sQry + "            Sqty = ISNULL(( ";
                                        sQry = sQry + "                               SELECT     COUNT(*)";
                                        sQry = sQry + "                               FROM       [ZPH_PY008] c";
                                        sQry = sQry + "                               WHERE      c.CLTCOD = A.U_CLTCOD";
                                        sQry = sQry + "                                          AND c.PosDate BETWEEN '" + YY + "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" + YY + "') + 1 ) + '0630'";
                                        sQry = sQry + "                                          AND c.MSTCOD  = a.Code";
                                        sQry = sQry + "                                          AND c.WorkType = 'D08'";
                                        sQry = sQry + "                                          AND c.PosDate <> '" + ymd + "'";
                                        sQry = sQry + "                          ),0) ";                                        //보전년차 사용수량
                                        sQry = sQry + "                    + ";
                                        sQry = sQry + "                    ISNULL(( ";
                                        sQry = sQry + "                               SELECT     COUNT(*) / 2.0";
                                        sQry = sQry + "                               FROM       [ZPH_PY008] c";
                                        sQry = sQry + "                               WHERE      c.CLTCOD = A.U_CLTCOD";
                                        sQry = sQry + "                                          AND c.PosDate BETWEEN '" + YY + "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" + YY + "') + 1 ) + '0630'";
                                        sQry = sQry + "                                          AND c.MSTCOD = a.Code";
                                        sQry = sQry + "                                          AND c.WorkType = 'D10'";
                                        sQry = sQry + "                                          AND c.PosDate <> '" + ymd + "'";
                                        sQry = sQry + "                           ),0)";                                        //보전반차 사용수량
                                        sQry = sQry + " FROM       [@PH_PY001A] a";
                                        sQry = sQry + " WHERE      a.U_CLTCOD = '" + CLTCOD + "'";
                                        sQry = sQry + "            AND a.Code = '" + MSTCOD + "'";

                                        oRecordSet01.DoQuery(sQry);

                                        if (WorkType == "D08")
                                        {
                                            JanQty = 1;
                                        }
                                        else if (WorkType == "D10")
                                        {
                                            JanQty = 0.5;
                                        }

                                        if (oRecordSet01.Fields.Item("Bqty").Value - oRecordSet01.Fields.Item("Sqty").Value < JanQty)
                                        {

                                            ErrNum = 2;
                                            oDS_PH_PY677B.SetValue("U_ColReg21", oRow - 1, "A00");
                                            oMat01.LoadFromDataSource();

                                            throw new Exception();

                                        }
                                        else
                                        {
                                            oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));      // 퇴근일자(P_OffDt)
                                            oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");                                            // 출근시각(P_GetTime)
                                            oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");                                            // 퇴근시각(P_OffTime)
                                            oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(1));                               // 교대일수
                                            PH_PY677_Time_ReSet(oRow);
                                            oMat01.LoadFromDataSource();
                                        }

                                        ProgressBar01.Value = 100;
                                        ProgressBar01.Stop();
                                        break;
                                    }
                            }
                        }
                        break;

                    case "MSTCOD":  //성명
                        oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'", "");
                        break;

                    case "ShiftDatCd": //근무형태
                        oForm.Items.Item("ShiftDatNm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code",  "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("ShiftDatCd").Specific.VALUE + "'", " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'");
                        break;

                    case "GNMUJOCd": //근무조
                        oForm.Items.Item("GNMUJONm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("GNMUJOCd").Specific.VALUE + "'", " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'");
                        break;

                    case "CLTCOD":

                        CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
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

                    case "TeamCode":

                        TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //담당콤보세팅
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "        SELECT      U_Code AS [Code],";
                        sQry = sQry + "             U_CodeNm As [Name]";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '2'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + "             AND U_Char1 = '" + TeamCode + "'";
                        sQry = sQry + " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "RspCode":

                        TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                        RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        // 반콤보세팅
                        oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "전체");
                        sQry = "        SELECT      U_Code AS [Code],";
                        sQry = sQry + "             U_CodeNm As [Name]";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '9'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + "             AND U_Char1 = '" + RspCode + "'";
                        sQry = sQry + "             AND U_Char2 = '" + TeamCode + "'";
                        sQry = sQry + " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
                        oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("연차휴가 잔여일수가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("보전휴가 잔여일수가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                //oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY677_Time_ReSet
        /// </summary>
        private void PH_PY677_Time_ReSet(int pRow)
        {
            try
            {
                oDS_PH_PY677B.SetValue("U_ColQty01", pRow - 1, Convert.ToString(0));  // 기본
                oDS_PH_PY677B.SetValue("U_ColQty02", pRow - 1, Convert.ToString(0));  // 연장
                oDS_PH_PY677B.SetValue("U_ColQty03", pRow - 1, Convert.ToString(0));  // 심야
                oDS_PH_PY677B.SetValue("U_ColQty06", pRow - 1, Convert.ToString(0));  // 조출
                oDS_PH_PY677B.SetValue("U_ColQty04", pRow - 1, Convert.ToString(0));  // 특근
                oDS_PH_PY677B.SetValue("U_ColQty05", pRow - 1, Convert.ToString(0));  // 특연
                oDS_PH_PY677B.SetValue("U_ColQty08", pRow - 1, Convert.ToString(0));  // 교육훈련
                oDS_PH_PY677B.SetValue("U_ColQty07", pRow - 1, Convert.ToString(0));  // 휴조
                oDS_PH_PY677B.SetValue("U_ColQty09", pRow - 1, Convert.ToString(0));  // 지각
                oDS_PH_PY677B.SetValue("U_ColQty10", pRow - 1, Convert.ToString(0));  // 조퇴
                oDS_PH_PY677B.SetValue("U_ColQty11", pRow - 1, Convert.ToString(0));  // 외출

                oDS_PH_PY677B.SetValue("U_ColTm02", pRow - 1, "00:00");               // 퇴근시각
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_Time_ReSet : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY677_Time_Calc_Main
        /// </summary>
        private object PH_PY677_Time_Calc_Main(string OffTime, int pRow)
        {
            object functionReturnValue = null;
            int i = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string ShiftDat = string.Empty; // 근무형태
            string GNMUJO = string.Empty;   // 근무조
            string DayOff = string.Empty;   // 평일휴일구분
            string PosDate = string.Empty;  // 기준일
            string GetDate = string.Empty;  // 출근일
            string OffDate = string.Empty;  // 퇴근일
            string GetTime = string.Empty;  // 출근시간
            string STime = string.Empty;
            string ETime = string.Empty;
            string FromTime = string.Empty;
            string ToTime = string.Empty;
            string NextDay = string.Empty;
            string TimeType = string.Empty;
            string WorkType = string.Empty;

            double hTime1 = 0;            // 오전10분휴식시간
            double hTime2 = 0;            // 점심휴식시간
            double hTime3 = 0;            // 오후 10분 휴식시간
            double hTime4 = 0;            // 저녁휴식시간
            double hTime5 = 0;            // 야간휴식시간
            double EarlyTo = 0;
            double SEarlyTo = 0;
            double Extend = 0;
            double SpExtend = 0;
            double Midnight = 0;
            double Base = 0;
            double Special = 0;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                ShiftDat = oDS_PH_PY677B.GetValue("U_ColReg07", pRow - 1).Trim();              // 근무형태
                GNMUJO =   oDS_PH_PY677B.GetValue("U_ColReg08", pRow - 1).Trim();              // 근무조

                PosDate = oDS_PH_PY677B.GetValue("U_ColDt05", pRow - 1).Trim();                // 일자
                GetDate = oDS_PH_PY677B.GetValue("U_ColDt01", pRow - 1).Trim();                // 출근일자
                OffDate = oDS_PH_PY677B.GetValue("U_ColDt02", pRow - 1).Trim();                // 퇴근일자
                DayOff  = oDS_PH_PY677B.GetValue("U_ColReg10", pRow - 1).Trim();               // 휴일 평일 구분
                GetTime = oDS_PH_PY677B.GetValue("U_ColTm01", pRow - 1).Trim();                // 출근시각
                OffTime = oDS_PH_PY677B.GetValue("U_ColTm02", pRow - 1).Trim();                // 퇴근시각

                //GetTime = "0" + GetTime;
                //GetTime = Strings.Right(GetTime, 4);
                GetTime = GetTime.PadLeft(4, '0');
                OffTime = OffTime.PadLeft(4, '0');

                WorkType = oDS_PH_PY677B.GetValue("U_ColReg21", pRow - 1).ToString().Trim();   // 근태구분

                sQry = "        SELECT   U_TimeType, ";
                sQry = sQry + "          U_FromTime, ";
                sQry = sQry + "          U_ToTime, ";
                sQry = sQry + "          U_NextDay ";
                sQry = sQry + " FROM     [@PH_PY002A] a ";
                sQry = sQry + "          INNER JOIN ";
                sQry = sQry + "          [@PH_PY002B] b ";
                sQry = sQry + "             On a.Code = b.Code ";
                sQry = sQry + " WHERE    a.U_CLTCOD = '" + CLTCOD + "'";                // 사업부
                sQry = sQry + "          AND a.U_SType = '" + ShiftDat + "'";           // 교대
                sQry = sQry + "          AND a.U_Shift = '" + GNMUJO + "'";             // 조
                sQry = sQry + "          AND b.U_DayType = '" + DayOff + "'";           // 평일

                oRecordSet.DoQuery(sQry);

                if ((oRecordSet.RecordCount == 0))
                {
                    //        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
                    ErrNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    FromTime = Convert.ToString(oRecordSet.Fields.Item(1).Value);
                    // FromTime = "0000" + FromTime;
                    // FromTime = Strings.Right(FromTime, 4);
                    FromTime = FromTime.PadLeft(4, '0');

                    ToTime = Convert.ToString(oRecordSet.Fields.Item(2).Value);
                    // ToTime = "0000" + ToTime;
                    // ToTime = Strings.Right(ToTime, 4);
                    ToTime = ToTime.PadLeft(4, '0');

                    NextDay = oRecordSet.Fields.Item(3).Value.Trim();
                    TimeType = oRecordSet.Fields.Item(0).Value.Trim();

                    if (string.IsNullOrEmpty(NextDay))
                    {
                        NextDay = "N";
                    }

                    if (NextDay == "N")
                    {
                        if (ToTime == "0000")
                        {
                            ToTime = "2400";
                        }
                    }

                    switch (TimeType)
                    {
                        case "40":
                            // 조출
                            // 출근당일이면
                            if (GetDate == PosDate)
                            {
                                // 출근시간 < 기준종료시간
                                if ( Convert.ToDouble(GetTime) < Convert.ToDouble(ToTime) )
                                {
                                    STime = GetTime;
                                    ETime = ToTime;

                                    if (DayOff == "1")
                                    {
                                        EarlyTo = PH_PY677_Time_Calc(STime, ETime);
                                    }
                                    else
                                    {
                                        SEarlyTo = PH_PY677_Time_Calc(STime, ETime);
                                    }

                                }
                            }
                            break;
                        case "10":
                        case "50":
                            ////정상근무시간
                            if ((GNMUJO == "11" | GNMUJO == "21"))
                            {
                                switch (NextDay)
                                {
                                    case "N":
                                        // 당일
                                        // 기준일자 = 출근일자
                                        if (PosDate == GetDate)
                                        {
                                            // 1교대1조, 2교대 1조당일
                                            // 출근시간 < 시작시간
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime; // 시작시간
                                            }
                                            else
                                            {
                                                STime = GetTime; // 출근시간
                                            }

                                            // 퇴근일이 틀릴때(다음날 퇴근일때)
                                            if (GetDate != OffDate)
                                            {
                                                ETime = ToTime; // 종료시간
                                            }
                                            else
                                            {
                                                // 퇴근시간 < 종료시간
                                                if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                                {
                                                    ETime = OffTime; // 퇴근시간
                                                }
                                                else
                                                {
                                                    ETime = ToTime; // 종료시간
                                                }
                                            }

                                            if (Convert.ToDouble(DayOff) == 1)
                                            {
                                                Base = Base + PH_PY677_Time_Calc(STime, ETime); // 평일
                                            }
                                            else
                                            {
                                                Special = Special + PH_PY677_Time_Calc(STime, ETime); // 휴일
                                            }
                                        }
                                        break;
                                    case "Y":
                                        // 익일
                                        break;
                                }
                            }
                            else
                            {
                                // 당일 익일 더해야 한다.
                                if (GNMUJO == "22")
                                {
                                    switch (NextDay)
                                    {
                                        case "N":
                                            // 당일
                                            // 기준일 = 출근일
                                            if (PosDate == GetDate)
                                            {
                                                // 출근시간 < 시작시간
                                                if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                                {
                                                    STime = FromTime; // 시작시간
                                                }
                                                else
                                                {
                                                    STime = GetTime; // 출근시간
                                                }

                                                // 퇴근일이 같으면(기준일 퇴근)
                                                if (GetDate == OffDate)
                                                {
                                                    // 퇴근시간 < 24시
                                                    if (Convert.ToDouble(OffTime) < 2400)
                                                    {
                                                        ETime = OffTime; // 퇴근시간
                                                    }
                                                    else
                                                    {
                                                        ETime = "2400";
                                                    }
                                                }
                                                else
                                                {
                                                    ETime = "2400";
                                                }

                                                if (Convert.ToDouble(DayOff) == 1)
                                                {
                                                    Base = Base + PH_PY677_Time_Calc(STime, ETime);
                                                }
                                                else
                                                {
                                                    Special = Special + PH_PY677_Time_Calc(STime, ETime);
                                                }
                                            }
                                            break;

                                        case "Y":
                                            // 익일
                                            // 기준일이 같으면 계산안함
                                            if (PosDate == OffDate)
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = "0000"; // 시작시간은 00시

                                                // 퇴근시간 < 종료시간
                                                if (Convert.ToDouble(OffTime ) < Convert.ToDouble(ToTime))
                                                {
                                                    ETime = OffTime; // 퇴근시간
                                                }
                                                else
                                                {
                                                    ETime = ToTime; // 종료시간
                                                }
                                            }

                                            if (Convert.ToDouble(DayOff) == 1)
                                            {
                                                Base = Base + PH_PY677_Time_Calc(STime, ETime);
                                            }
                                            else
                                            {
                                                Special = Special + PH_PY677_Time_Calc(STime, ETime);
                                            }
                                            break;
                                    }
                                }
                            }
                            break;
                        case "65":
                        case "66":
                        case "15":
                            // 오전, 오후 휴식시간, 점심시간
                            // 1교대1조 2교대 1조
                            if ((GNMUJO == "11" | GNMUJO == "21" | GNMUJO == "22"))
                            {
                                switch (NextDay)
                                {
                                    case "N":
                                        // 당일
                                        // 당일 출근이 아니면
                                        if (PosDate != GetDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            // 당일퇴근
                                            if (PosDate == OffDate)
                                            {
                                                // 출근시간 < 시작시간
                                                if (Convert.ToDouble(GetTime)  < Convert.ToDouble(FromTime))
                                                {
                                                    // 시작시간 <= 퇴근시간
                                                    if (Convert.ToDouble(FromTime) <= Convert.ToDouble(OffTime))
                                                    {
                                                        STime = FromTime;
                                                        // 시작시간
                                                        // 종료시간 < 퇴근시간
                                                        if (Convert.ToDouble(ToTime) < Convert.ToDouble(OffTime))
                                                        {
                                                            ETime = ToTime;
                                                        }
                                                        else
                                                        {
                                                            ETime = OffTime;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        STime = "0000";
                                                        ETime = "0000";
                                                    }
                                                }
                                                else
                                                {
                                                    // 출근시간 > 종료시간
                                                    if (Convert.ToDouble(GetTime) > Convert.ToDouble(ToTime))
                                                    {
                                                        STime = "0000";
                                                        ETime = "0000";
                                                    }
                                                    else
                                                    {
                                                        STime = GetTime;
                                                        // 출근시간
                                                        // 종료시간 < 퇴근시간
                                                        if (Convert.ToDouble(ToTime) < Convert.ToDouble(OffTime))
                                                        {
                                                            ETime = ToTime; // 종료시간
                                                        }
                                                        else
                                                        {
                                                            ETime = OffTime; // 퇴근시간
                                                        }
                                                    }
                                                }
                                                // 다음날 퇴근
                                            }
                                            else
                                            {
                                                // 종료시간 < 출근시간
                                                if (Convert.ToDouble(ToTime) < Convert.ToDouble(GetTime))
                                                {
                                                    STime = "0000";
                                                    ETime = "0000";
                                                }
                                                else
                                                {
                                                    // 출근시간 < 시작시간
                                                    if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                                    {
                                                        STime = FromTime;
                                                        ETime = ToTime;

                                                        // 시작시간이후 출근
                                                    }
                                                    else
                                                    {
                                                        STime = GetTime;
                                                    }
                                                    ETime = ToTime;
                                                }
                                            }
                                        }
                                        break;
                                    case "Y":
                                        // 익일
                                        // 기준일 = 퇴근일(당일퇴근)
                                        if (PosDate == OffDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            // 퇴근시간 < 시작시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = FromTime;
                                                // 퇴근시간 < 종료시간
                                                if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                                {
                                                    ETime = OffTime;
                                                }
                                                else
                                                {
                                                    ETime = ToTime;
                                                }
                                            }
                                        }
                                        break;
                                }

                            }

                            hTime1 = PH_PY677_Time_Calc(STime, ETime); // 오전휴식시간

                            // 야간조(2교대) 오전휴식일경우
                            if (GNMUJO == "22" & TimeType == "65")
                            {
                                Midnight = Midnight - hTime1;
                            }
                            else
                            {
                                if (DayOff == "1")
                                {
                                    Base = Base - hTime1;
                                }
                                else
                                {
                                    Special = Special - hTime1;
                                }
                            }
                            break;

                        case "20":
                        case "60":
                            // 연장근무
                            // 1교대1조 2교대 1조
                            if ((GNMUJO == "11" | GNMUJO == "21"))
                            {
                                switch (NextDay)
                                {
                                    case "N":
                                        // 당일
                                        //  출근일 <> 퇴근일
                                        if (PosDate != OffDate)
                                        {
                                            // 출근시간 < 시작시간
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime;
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }
                                            ETime = "2400";// 당일 퇴근
                                        }
                                        else
                                        {
                                            // 퇴근시간 < 시작시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = FromTime; // 종료시간
                                                ETime = OffTime;  // 퇴근시간
                                            }
                                        }
                                        if (DayOff == "1")
                                        {
                                            Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        break;
                                    case "Y":
                                        // 익일
                                        // 기준일 <> 퇴근일
                                        if (PosDate != OffDate)
                                        {
                                            STime = "0000";
                                            // 퇴근시간 < 종료시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }

                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        if (DayOff == "1")
                                        {
                                            Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        break;

                                }
                                // 2교대 2조
                            }
                            else if (GNMUJO == "22")
                            {
                                switch (NextDay)
                                {
                                    case "N":
                                        //당일
                                        break;
                                    // 연장근무없다
                                    case "Y":
                                        // 익일
                                        //  기준일 <> 퇴근일
                                        if (PosDate != OffDate)
                                        {
                                            // 퇴근시간 < 시작시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = FromTime;
                                                ETime = OffTime;

                                            }
                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }

                                        if (DayOff == "1")
                                        {
                                            Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
                                        }
                                        break;
                                }
                            }
                            break;
                        case "25":
                            // 저녁휴식
                            break;
                        case "30":
                            // 심야시간
                            // 1교대1조 2교대 1조
                            if ((GNMUJO == "11" | GNMUJO == "21" | GNMUJO == "22"))
                            {
                                switch (NextDay)
                                {
                                    case "N":
                                        // 당일
                                        // 기준일자 <> 퇴근일자
                                        if (PosDate != OffDate)
                                        {
                                            // 출근시간 < 시작시간
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime; // 시작시간
                                            }
                                            else
                                            {
                                                STime = GetTime; // 출근시간
                                            }
                                            ETime = "2400"; // 당일퇴근시
                                        }
                                        else
                                        {
                                            //퇴근시간 < 시작시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                // 출근시간 < 시작시간
                                                if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                                {
                                                    STime = FromTime; // 시작시간
                                                }
                                                else
                                                {
                                                    STime = GetTime;
                                                }

                                                // 퇴근시간 < 종료시간
                                                if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                                {
                                                    ETime = OffTime;
                                                }
                                                else
                                                {
                                                    ETime = ToTime;// 종료시간
                                                }
                                            }
                                        }
                                        Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime);
                                        break;
                                    case "Y":
                                        // 익일
                                        // 출근일 <> 퇴근일
                                        if (PosDate != OffDate)
                                        {
                                            STime = "0000";
                                            // 퇴근시간 < 종료시간
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime; // 퇴근시간
                                            }
                                            else
                                            {
                                                ETime = ToTime; // 종료시간
                                            }
                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime);
                                        break;
                                }

                            }
                            break;
                        case "35":
                            // 야간휴식 수정해야함
                            // hTime5
                            switch (NextDay)
                            {
                                case "N":
                                    // 당일
                                    // 기준일자 = 퇴근일자
                                    if (PosDate == OffDate)
                                    {
                                        // 퇴근시간 < 시작시간
                                        if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime;
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }

                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // 다음날 퇴근
                                        // 출근시간 < 시작시간
                                        if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = FromTime;
                                        }
                                        else
                                        {
                                            STime = GetTime;
                                        }

                                        ETime = ToTime;
                                    }
                                    break;

                                case "Y":
                                    // 기준일자 = 퇴근일자
                                    if (PosDate == OffDate)
                                    {
                                        STime = "0000";
                                        ETime = "0000";
                                    }
                                    else
                                    {
                                        // 퇴근시간 < 시작시간
                                        if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            STime = FromTime;
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }
                                        }
                                    }
                                    break;
                            }

                            hTime5 = PH_PY677_Time_Calc(STime, ETime);

                            // 평일
                            if (DayOff == "1")
                            {
                                // 2교대 2조(야간조)
                                if (GNMUJO == "22")
                                {
                                    Base = Base - hTime5; // 기본근무
                                    Midnight = Midnight - hTime5;  // 심야시간에서 차감
                                    // 그외 주간조
                                }
                                else
                                {
                                    Extend = Extend - hTime5;
                                    // 연장근무에서 차감
                                    Midnight = Midnight - hTime5;
                                    // 심야시간에서 차감
                                }
                            }
                            else
                            {
                                // 휴일
                                if (GNMUJO == "22")
                                {
                                    Special = Special - hTime5;
                                    Midnight = Midnight - hTime5; // 심야시간에서 차감
                                }
                                else
                                {
                                    // 다음날 퇴근은 연장근무임
                                    SpExtend = SpExtend - hTime5; // 연장근무에서 차감
                                    Midnight = Midnight - hTime5; // 심야시간에서 차감
                                }
                            }
                            break;

                    }
                    oRecordSet.MoveNext();
                }

                oDS_PH_PY677B.SetValue("U_ColQty06", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(EarlyTo)));                //조출
                oDS_PH_PY677B.SetValue("U_ColQty01", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Base, WorkType)));         //기본
                oDS_PH_PY677B.SetValue("U_ColQty02", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Extend)));                 //연장
                oDS_PH_PY677B.SetValue("U_ColQty03", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Midnight)));               //심야

                oDS_PH_PY677B.SetValue("U_ColQty07", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(SEarlyTo)));               //휴조
                oDS_PH_PY677B.SetValue("U_ColQty04", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Special)));                //특근
                oDS_PH_PY677B.SetValue("U_ColQty05", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(SpExtend)));               //특근연장

                return functionReturnValue;
                
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근태시간구분등록 자료가 없습니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_Time_Calc_Main_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY677_Time_Calc
        /// </summary>
        /// <param name="GetTime"></param>
        /// <param name="OffTime"></param>
        /// <returns></returns>
        private double PH_PY677_Time_Calc(string GetTime, string OffTime)
        {
            double functionReturnValue = 0;
            int hh, STime, ETime = 0;
            double MM, Ret = 0;
    
            try
                {
                    STime = Convert.ToInt32(Convert.ToDouble( GetTime.Substring(0, 2) ) ) * 3600 + Convert.ToInt32(Convert.ToDouble(GetTime.Substring(2, 2))) * 60;
                    ETime = Convert.ToInt32(Convert.ToDouble( OffTime.Substring(0, 2) ) ) * 3600 + Convert.ToInt32(Convert.ToDouble(OffTime.Substring(2, 2))) * 60;
                    Ret = ETime - STime;

                    functionReturnValue = Ret;
                    return functionReturnValue;
                }
           
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_Time_Calc_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return functionReturnValue;
            }
            finally
            {
                
            }
        }

        /// <summary>
        /// PH_PY677_hhmm_Calc
        /// </summary>
        /// <param name="mTime"></param>
        /// <param name="pWorkType"></param>
        /// <returns></returns>
        private double PH_PY677_hhmm_Calc(double mTime, string pWorkType = "")
        {
            double functionReturnValue = 0;
            int hh = 0;
            double MM = 0;
            double Ret = 0;

            try
            {
                hh = Convert.ToInt32(Math.Truncate(mTime / 3600));
                MM = ((mTime) % 3600) / 60;

                if (MM > 0)
                {
                    if (MM > 30)
                    {
                        MM = 1;
                    }
                    else
                    {
                        MM = 0.5;
                    }
                }
                else if (MM == 0)
                {
                    MM = 0;
                }
                else
                {
                    if (MM < -30)
                    {
                        MM = -1;
                    }
                    else
                    {
                        MM = -0.5;
                    }
                }

                Ret = hh + MM;

                if (pWorkType == "D09" | pWorkType == "D10")
                {
                    Ret = 4;
                }

                functionReturnValue = Ret;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_hhmm_Calc_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return functionReturnValue;
            }
            finally
            {
            }
        }


        /// <summary>
        /// Matrix 체크박스 전체 선택
        /// </summary>
        private void PH_PY677_CheckAll()
        {
            string CheckType = string.Empty;
            int loopCount = 0;

            CheckType = "Y";

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY677B.GetValue("U_ColReg20", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break; 
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY677B.Offset = loopCount;
                    if (CheckType == "N")
                    {
                        oDS_PH_PY677B.SetValue("U_ColReg20", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY677B.SetValue("U_ColReg20", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_CheckAll_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

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
                    if (pVal.ItemUID == "PH_PY677")
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

                    // 수정 버튼
                    if (pVal.ItemUID == "BtnUpdate")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PH_PY677_UpdateData();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY677_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    // 조회
                    }
                    else if (pVal.ItemUID == "BtnSearch")
                    {
                        PH_PY677_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnChk")
                    {
                        PH_PY677_CheckAll();
                    }
                    else if (pVal.ItemUID == "BtnConfirm")
                    {
                        PH_PY677_ConfirmAll();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY677")
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
        /// PH_PY677_ConfirmAll
        /// </summary>
        private void PH_PY677_ConfirmAll()
        {
            string ConfirmType = string.Empty;
            short loopCount = 0;
            ConfirmType = "Y";

            try
            {
                oForm.Freeze(true);
                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY677B.GetValue("U_ColReg16", loopCount).ToString().Trim() == "N")
                    {
                        ConfirmType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY677B.Offset = loopCount;
                    if (ConfirmType == "N")
                    {
                        oDS_PH_PY677B.SetValue("U_ColReg16", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY677B.SetValue("U_ColReg16", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY677_ConfirmAll_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", ""); //사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ShiftDatCd", ""); //근무형태
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "GNMUJOCd", ""); //근무조
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
                        if (pVal.ColUID == "P_WorkType")
                        {
                            PH_PY677_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                    }
                    else
                    {
                        PH_PY677_FlushToItemValue(pVal.ItemUID, 0, "");
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oMat01.SelectRow(pVal.Row, true, false);

                    }
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
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "P_GetTime" || pVal.ColUID == "P_OffTime")
                            {
                                PH_PY677_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            oMat01.AutoResizeColumns();
                        }
                        else
                        {
                            PH_PY677_FlushToItemValue(pVal.ItemUID, 0, "");
                        }

                        oForm.Update();
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
                oForm.Freeze(false);
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
                    PH_PY677_FormItemEnabled();
                    //PH_PY677_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY677B);
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
                    PH_PY677_FormResize();
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    //            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY677A", "U_CardCode,U_CardName")
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
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기
                            PH_PY677_Add_MatrixRow(oMat01.VisualRowCount, false); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
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
                            //엑셀 내보내기 이후 처리_S
                            oForm.Freeze(true);
                            oDS_PH_PY677B.RemoveRecord(oDS_PH_PY677B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            //엑셀 내보내기 이후 처리_E
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

        /// <summary>
        /// ROW_DELETE(Raise_FormMenuEvent에서 호출)
        /// 해당 클래스에서는 사용되지 않음
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pval"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        {
            // ERROR: Not supported in C#: OnErrorStatement

            int i = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pval.BeforeAction == true)
                    {
                        //            If (PH_PY677_Validate("행삭제") = False) Then
                        //                BubbleEvent = False
                        //                Exit Sub
                        //            End If
                        ////행삭제전 행삭제가능여부검사
                    }
                    else if (pval.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PH_PY677B.RemoveRecord(oDS_PH_PY677B.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY677_Add_MatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY677B.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY677_Add_MatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #region 구 이벤트 소스코드, 최종테스트 후 삭제 요망

        #region Raise_FormMenuEvent
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {

        //			}
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
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
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
        //			return;
        //			Raise_FormDataEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}

        //			return;
        //			Raise_RightClickEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormItemEvent
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //				
        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {


        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {



        //			}

        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

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
//	internal class PH_PY677
//	{
////****************************************************************************************************************
//////  File               : PH_PY677.cls
//////  Module             : 인사관리>근태관리>근태리포트
//////  Desc               : 일일근태이상자조회
//////  FormType           : PH_PY677
//////  Create Date(Start) : 2013.03.25
//////  Create Date(End)   :
//////  Creator            : Song Myoung gyu
//////  Modified Date      :
//////  Modifier           :
//////  Company            : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm;
//		public SAPbouiCOM.Matrix oMat01;
//			//등록헤더
//		private SAPbouiCOM.DBDataSource oDS_PH_PY677A;
//			//등록라인
//		private SAPbouiCOM.DBDataSource oDS_PH_PY677B;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string oLastItemUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string oLastColUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oLastColRow01;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(string oFromDocEntry01 = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			string oInnerXml = null;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY677.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PH_PY677_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID01, "PH_PY677");
//			////폼추가
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID01);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

//			oForm.Freeze(true);
//			PH_PY677_CreateItems();
//			PH_PY677_ComboBox_Setting();
//			PH_PY677_CF_ChooseFromList();
//			PH_PY677_EnableMenus();
//			PH_PY677_SetDocument(oFromDocEntry01);
//			PH_PY677_FormResize();

//			//    Call PH_PY677_Add_MatrixRow(0, True)
//			//    Call PH_PY677_LoadCaption
//			PH_PY677_FormItemEnabled();

//			oForm.EnableMenu(("1283"), false);
//			//// 삭제
//			oForm.EnableMenu(("1286"), false);
//			//// 닫기
//			oForm.EnableMenu(("1287"), false);
//			//// 복제
//			oForm.EnableMenu(("1285"), false);
//			//// 복원
//			oForm.EnableMenu(("1284"), false);
//			//// 취소
//			oForm.EnableMenu(("1293"), false);
//			//// 행삭제
//			oForm.EnableMenu(("1281"), false);
//			oForm.EnableMenu(("1282"), true);

//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY677A]"
//			//    Call RecordSet01.DoQuery(sQry)
//			//    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
//			//        Call oDS_PH_PY677A.setValue("DocEntry", 0, 1)
//			//    Else
//			//        Call oDS_PH_PY677A.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
//			//    End If
//			//
//			//    Call PH_PY677_FormReset '폼초기화 추가(2013.01.29 송명규)

//			oForm.Update();
//			oForm.Freeze(false);

//			//기간(월)
//			//UPGRADE_WARNING: oForm.Items(FrDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FrDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM01");
//			//UPGRADE_WARNING: oForm.Items(ToDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ToDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
//			//사번 포커스
//			oForm.Items.Item("MSTCOD").Click();

//			//UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("BtnChk").Specific.Caption = "전체선택";

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;

//			return;
//			LoadForm_Error:
//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_LoadCaption()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY677_LoadCaption()
//			//해당모듈    : PH_PY677
//			//기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				//UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
//				oForm.Items.Item("BtnDelete").Enabled = false;
//				//    ElseIf oForm.Mode = fm_OK_MODE Then
//				//        oForm.Items("BtnAdd").Specific.Caption = "확인"
//			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//				//UPGRADE_WARNING: oForm.Items().Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
//				oForm.Items.Item("BtnDelete").Enabled = true;
//			}

//			oForm.Freeze(false);

//			return;
//			PH_PY677_LoadCaption_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "Delete_EmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void PH_PY677_Add_MatrixRow(int oRow, ref bool RowIserted = false)
//		{
//			//******************************************************************************
//			//Function ID : PH_PY677_Add_MatrixRow
//			//해당모듈    : PH_PY677
//			//기능        : 메트릭스 행 추가
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			////행추가여부
//			if (RowIserted == false) {
//				oDS_PH_PY677B.InsertRecord((oRow));
//			}

//			oMat01.AddRow();
//			oDS_PH_PY677B.Offset = oRow;
//			oDS_PH_PY677B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

//			oMat01.LoadFromDataSource();
//			return;
//			PH_PY677_Add_MatrixRow_Error:

//			MDC_Com.MDC_GF_Message(ref "PH_PY677_Add_MatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

//		public void PH_PY677_MTX01()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY677_MTX01()
//			//해당모듈    : PH_PY677
//			//기능        : 데이터 조회
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;
//			short ErrNum = 0;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string CLTCOD = null;
//			//사업장
//			string FrDate = null;
//			//시작일자
//			string ToDate = null;
//			//종료일자
//			string TeamCode = null;
//			//부서
//			string RspCode = null;
//			//담당
//			string ClsCode = null;
//			//반
//			string ShiftDat = null;
//			//근무형태
//			string GNMUJO = null;
//			//근무조
//			string MSTCOD = null;
//			//사원번호
//			//UPGRADE_NOTE: Class이(가) Class_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string Class_Renamed = null;
//			//기찰이상구분(2014.04.10 송명규 추가)
//			string Confirm = null;
//			//근태기찰정상확인(2013.03.29 송명규 추가)
//			string WorkType = null;
//			//근태구분(2014.05.13 송명규 추가)

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FrDate = Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE);
//			//시작일자
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ToDate = Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE);
//			//종료일자
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.VALUE);
//			//부서
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			RspCode = Strings.Trim(oForm.Items.Item("RspCode").Specific.VALUE);
//			//담당
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ClsCode = Strings.Trim(oForm.Items.Item("ClsCode").Specific.VALUE);
//			//반
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ShiftDat = Strings.Trim(oForm.Items.Item("ShiftDatCd").Specific.VALUE);
//			//근무형태
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			GNMUJO = Strings.Trim(oForm.Items.Item("GNMUJOCd").Specific.VALUE);
//			//근무조
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
//			//사원번호
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Class_Renamed = Strings.Trim(oForm.Items.Item("Class").Specific.VALUE);
//			//기찰이상구분
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Confirm = Strings.Trim(oForm.Items.Item("Confirm").Specific.VALUE);
//			//근태기찰정상확인
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			WorkType = Strings.Trim(oForm.Items.Item("WorkType").Specific.VALUE);
//			//근태구분

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//			oForm.Freeze(true);

//			sQry = "            EXEC [PH_PY677_01] ";
//			sQry = sQry + "'" + CLTCOD + "',";
//			//사업장
//			sQry = sQry + "'" + FrDate + "',";
//			//시작일자
//			sQry = sQry + "'" + ToDate + "',";
//			//종료일자
//			sQry = sQry + "'" + TeamCode + "',";
//			//부서
//			sQry = sQry + "'" + RspCode + "',";
//			//담당
//			sQry = sQry + "'" + ClsCode + "',";
//			//반
//			sQry = sQry + "'" + ShiftDat + "',";
//			//근무형태
//			sQry = sQry + "'" + GNMUJO + "',";
//			//근무조
//			sQry = sQry + "'" + MSTCOD + "',";
//			//사원번호
//			sQry = sQry + "'" + Class_Renamed + "',";
//			//기찰이상구분
//			sQry = sQry + "'" + Confirm + "',";
//			//근태기찰정상확인
//			sQry = sQry + "'" + WorkType + "'";
//			//근태구분

//			oRecordSet01.DoQuery(sQry);

//			oMat01.Clear();
//			oDS_PH_PY677B.Clear();
//			oMat01.FlushToDataSource();
//			oMat01.LoadFromDataSource();

//			if ((oRecordSet01.RecordCount == 0)) {

//				ErrNum = 1;

//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//				//        Call PH_PY677_Add_MatrixRow(0, True)
//				//        Call PH_PY677_LoadCaption

//				goto PH_PY677_MTX01_Error;

//				return;
//			}

//			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//				if (i + 1 > oDS_PH_PY677B.Size) {
//					oDS_PH_PY677B.InsertRecord((i));
//				}

//				oMat01.AddRow();
//				oDS_PH_PY677B.Offset = i;

//				oDS_PH_PY677B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PH_PY677B.SetValue("U_ColReg20", i, Strings.Trim(oRecordSet01.Fields.Item("Chk").Value));
//				//선택
//				oDS_PH_PY677B.SetValue("U_ColDt05", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("PosDate").Value), "YYYYMMDD"));
//				//일자
//				oDS_PH_PY677B.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("TeamName").Value));
//				//부서
//				oDS_PH_PY677B.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("RspName").Value));
//				//담당
//				oDS_PH_PY677B.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("ClsName").Value));
//				//반
//				oDS_PH_PY677B.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("MSTCOD").Value));
//				//사번
//				oDS_PH_PY677B.SetValue("U_ColReg06", i, Strings.Trim(oRecordSet01.Fields.Item("MSTNAM").Value));
//				//성명
//				oDS_PH_PY677B.SetValue("U_ColReg07", i, Strings.Trim(oRecordSet01.Fields.Item("ShiftDat").Value));
//				//근무형태
//				oDS_PH_PY677B.SetValue("U_ColReg08", i, Strings.Trim(oRecordSet01.Fields.Item("GNMUJO").Value));
//				//근무조
//				oDS_PH_PY677B.SetValue("U_ColReg09", i, Strings.Trim(oRecordSet01.Fields.Item("DayWeek").Value));
//				//요일
//				oDS_PH_PY677B.SetValue("U_ColReg10", i, Strings.Trim(oRecordSet01.Fields.Item("DayType").Value));
//				//요일구분
//				oDS_PH_PY677B.SetValue("U_ColDt01", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("P_GetDt").Value), "YYYYMMDD"));
//				//출근일자(계획)
//				oDS_PH_PY677B.SetValue("U_ColTm01", i, Strings.Trim(oRecordSet01.Fields.Item("P_GetTime").Value));
//				//출근시각(계획)
//				oDS_PH_PY677B.SetValue("U_ColDt02", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("P_OffDt").Value), "YYYYMMDD"));
//				//퇴근일자(계획)
//				oDS_PH_PY677B.SetValue("U_ColTm02", i, Strings.Trim(oRecordSet01.Fields.Item("P_OffTime").Value));
//				//퇴근시각(계획)
//				oDS_PH_PY677B.SetValue("U_ColQty01", i, Strings.Trim(oRecordSet01.Fields.Item("P_Base").Value));
//				//기본(계획)
//				oDS_PH_PY677B.SetValue("U_ColQty02", i, Strings.Trim(oRecordSet01.Fields.Item("P_Extend").Value));
//				//연장(계획)
//				oDS_PH_PY677B.SetValue("U_ColQty03", i, Strings.Trim(oRecordSet01.Fields.Item("P_MidNt").Value));
//				//심야
//				oDS_PH_PY677B.SetValue("U_ColQty04", i, Strings.Trim(oRecordSet01.Fields.Item("P_Special").Value));
//				//특근(계획)
//				oDS_PH_PY677B.SetValue("U_ColQty05", i, Strings.Trim(oRecordSet01.Fields.Item("P_SpExtend").Value));
//				//특연(계획)
//				oDS_PH_PY677B.SetValue("U_ColQty06", i, Strings.Trim(oRecordSet01.Fields.Item("P_EarlyTo").Value));
//				//조출
//				oDS_PH_PY677B.SetValue("U_ColQty07", i, Strings.Trim(oRecordSet01.Fields.Item("P_SEarlyTo").Value));
//				//휴조
//				oDS_PH_PY677B.SetValue("U_ColQty08", i, Strings.Trim(oRecordSet01.Fields.Item("P_EducTran").Value));
//				//교육훈련
//				oDS_PH_PY677B.SetValue("U_ColQty09", i, Strings.Trim(oRecordSet01.Fields.Item("P_LateTo").Value));
//				//지각
//				oDS_PH_PY677B.SetValue("U_ColQty10", i, Strings.Trim(oRecordSet01.Fields.Item("P_EarlyOff").Value));
//				//조퇴
//				oDS_PH_PY677B.SetValue("U_ColReg21", i, Strings.Trim(oRecordSet01.Fields.Item("P_WorkType").Value));
//				//근태구분
//				oDS_PH_PY677B.SetValue("U_ColReg22", i, Strings.Trim(oRecordSet01.Fields.Item("P_Comment").Value));
//				//비고
//				oDS_PH_PY677B.SetValue("U_ColQty11", i, Strings.Trim(oRecordSet01.Fields.Item("P_GoOut").Value));
//				//외출
//				oDS_PH_PY677B.SetValue("U_ColReg16", i, Strings.Trim(oRecordSet01.Fields.Item("P_Confirm").Value));
//				//기찰정상확인
//				oDS_PH_PY677B.SetValue("U_ColReg17", i, Strings.Trim(oRecordSet01.Fields.Item("R_GetDt").Value));
//				//출근일자(기찰)
//				oDS_PH_PY677B.SetValue("U_ColTm03", i, Strings.Trim(oRecordSet01.Fields.Item("R_GetTime").Value));
//				//출근시각(기찰)
//				oDS_PH_PY677B.SetValue("U_ColReg19", i, Strings.Trim(oRecordSet01.Fields.Item("R_OffDt").Value));
//				//퇴근일자(기찰)
//				oDS_PH_PY677B.SetValue("U_ColTm04", i, Strings.Trim(oRecordSet01.Fields.Item("R_OffTime").Value));
//				//퇴근시각(기찰)
//				//        Call oDS_PH_PY677B.setValue("U_ColQty05", i, Trim(oRecordSet01.Fields("R_Base").VALUE)) '기본(기찰)
//				//        Call oDS_PH_PY677B.setValue("U_ColQty06", i, Trim(oRecordSet01.Fields("R_Extend").VALUE)) '연장(기찰)
//				//        Call oDS_PH_PY677B.setValue("U_ColQty07", i, Trim(oRecordSet01.Fields("R_Special").VALUE)) '특근(기찰)
//				//        Call oDS_PH_PY677B.setValue("U_ColQty08", i, Trim(oRecordSet01.Fields("R_SpExtend").VALUE)) '특연(기찰)
//				//        Call oDS_PH_PY677B.setValue("U_ColQty09", i, Trim(oRecordSet01.Fields("R_TotTime").VALUE)) '총근무시간(기찰)
//				oDS_PH_PY677B.SetValue("U_ColQty12", i, Strings.Trim(oRecordSet01.Fields.Item("Rotation").Value));
//				//교대일수
//				oDS_PH_PY677B.SetValue("U_ColReg24", i, Strings.Trim(oRecordSet01.Fields.Item("R_YN").Value));
//				//기찰완료여부
//				oDS_PH_PY677B.SetValue("U_ColReg25", i, Strings.Trim(oRecordSet01.Fields.Item("WkAbCls").Value));
//				//근태이상분류
//				oDS_PH_PY677B.SetValue("U_ColReg26", i, Strings.Trim(oRecordSet01.Fields.Item("WkAbCmt").Value));
//				//근태이상사유
//				oDS_PH_PY677B.SetValue("U_ColReg27", i, Strings.Trim(oRecordSet01.Fields.Item("ActText").Value));
//				//근무내용
//				oDS_PH_PY677B.SetValue("U_ColReg28", i, Strings.Trim(oRecordSet01.Fields.Item("RotateYN").Value));
//				//교대인정

//				oRecordSet01.MoveNext();
//				ProgBar01.Value = ProgBar01.Value + 1;
//				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

//			}

//			oMat01.LoadFromDataSource();
//			oMat01.AutoResizeColumns();
//			ProgBar01.Stop();
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			return;
//			PH_PY677_MTX01_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//    ProgBar01.Stop
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY677_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		public void PH_PY677_DeleteData()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY677_DeleteData()
//			//해당모듈    : PH_PY677
//			//기능        : 기본정보 삭제
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;
//			short ErrNum = 0;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string DocEntry = null;

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {

//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);

//				sQry = "SELECT COUNT(*) FROM [@PH_PY677A] WHERE DocEntry = '" + DocEntry + "'";
//				oRecordSet01.DoQuery(sQry);

//				if ((oRecordSet01.RecordCount == 0)) {
//					ErrNum = 1;
//					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//					goto PH_PY677_DeleteData_Error;
//				} else {
//					sQry = "EXEC PH_PY677_04 '" + DocEntry + "'";
//					oRecordSet01.DoQuery(sQry);
//				}
//			}

//			MDC_Com.MDC_GF_Message(ref "삭제 완료!", ref "S");

//			//    Call PH_PY677_FormReset

//			//    oForm.Mode = fm_ADD_MODE

//			//    Call oForm.Items("BtnSearch").Click(ct_Regular)

//			//    oMat01.Clear
//			//    oMat01.FlushToDataSource
//			//    oMat01.LoadFromDataSource
//			//    Call PH_PY677_Add_MatrixRow(0, True)

//			return;
//			PH_PY677_DeleteData_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "삭제대상이 없습니다. 확인하세요.", ref "W");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY677_DeleteData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		public bool PH_PY677_UpdateData()
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : PH_PY677_UpdateData()
//			//해당모듈    : PH_PY677
//			//기능        : 기본정보 수정
//			//인수        : 없음
//			//반환값      : 없음
//			//특이사항    : 1. 근무형태, 근무조 추가(2017.09.05 송명규)
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			short j = 0;
//			string sQry = null;

//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string CLTCOD = null;
//			//사업장
//			string PosDate = null;
//			//일자
//			string MSTCOD = null;
//			//사번
//			string JIGTYP = null;
//			//지급타입
//			string PAYTYP = null;
//			//급여타입
//			string JIGCOD = null;
//			//직급코드
//			string ShiftDat = null;
//			//근무형태
//			string GNMUJO = null;
//			//근무조
//			string P_WorkType = null;
//			//근태구분
//			string P_Confirm = null;
//			//기찰정상확인
//			string P_GetTime = null;
//			//출근시각
//			string P_OffDt = null;
//			//퇴근일자
//			string P_OffTime = null;
//			//퇴근시각
//			double P_Base = 0;
//			//기본
//			double P_Extend = 0;
//			//연장
//			double P_Special = 0;
//			//특근
//			double P_SpExtend = 0;
//			//특연
//			double P_Midnight = 0;
//			//심야
//			double P_EarlyTo = 0;
//			//조출
//			double P_SEarlyTo = 0;
//			//휴조
//			double P_EducTran = 0;
//			//교육훈련
//			double P_LateTo = 0;
//			//지각
//			double P_EarlyOff = 0;
//			//조퇴
//			double P_GoOut = 0;
//			//외출
//			string P_Comment = null;
//			//비고
//			string DangerCd = null;
//			//비고
//			string WkAbCls = null;
//			//근태이상분류
//			string WkAbCmt = null;
//			//근태이상사유
//			string ActText = null;
//			//근무내용
//			string RotateYN = null;
//			//교대인정

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, false);

//			oMat01.FlushToDataSource();
//			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
//				if (Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg20", i)) == "Y") {

//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//					//사업장
//					PosDate = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", i));
//					//일자
//					MSTCOD = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg05", i));
//					//사번
//					ShiftDat = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg07", i));
//					//근무형태
//					GNMUJO = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg08", i));
//					//근무조
//					P_WorkType = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg21", i));
//					//근태구분
//					P_Confirm = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg16", i));
//					//기찰정상확인
//					P_GetTime = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColTm01", i));
//					//출근시각
//					P_OffDt = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt02", i));
//					//퇴근일자
//					P_OffTime = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColTm02", i));
//					//퇴근시각
//					P_Base = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty01", i)));
//					//기본
//					P_Extend = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty02", i)));
//					//연장
//					P_Special = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty04", i)));
//					//특근
//					P_SpExtend = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty05", i)));
//					//특연
//					P_Midnight = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty03", i)));
//					//심야
//					P_EarlyTo = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty06", i)));
//					//조출
//					P_SEarlyTo = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty07", i)));
//					//휴조
//					P_EducTran = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty08", i)));
//					//교육훈련
//					P_LateTo = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty09", i)));
//					//지각
//					P_EarlyOff = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty10", i)));
//					//조퇴
//					P_GoOut = Convert.ToDouble(Strings.Trim(oDS_PH_PY677B.GetValue("U_ColQty11", i)));
//					//외출
//					P_Comment = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg22", i));
//					//비고
//					WkAbCls = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg25", i));
//					//근태이상분류
//					WkAbCmt = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg26", i));
//					//근태이상사유
//					ActText = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg27", i));
//					//근무내용
//					RotateYN = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg28", i));
//					//교대인정
//					//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					DangerCd = MDC_GetData.Get_ReData(ref "DangerCD", ref "PosDate", ref "ZPH_PY008", ref "'" + Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", i)) + "'", ref " and mstcod ='" + MSTCOD + "'");

//					sQry = "Select   U_JIGTYP";
//					sQry = sQry + ", U_PAYTYP";
//					sQry = sQry + ", U_JIGCOD ";
//					sQry = sQry + "from [@PH_PY001A] ";
//					sQry = sQry + "where u_status <> '5' and code ='" + MSTCOD + "' ";

//					RecordSet01.DoQuery(sQry);

//					JIGTYP = Convert.ToString(RecordSet01.Fields.Item(0).Value);
//					PAYTYP = Convert.ToString(RecordSet01.Fields.Item(1).Value);
//					JIGCOD = Convert.ToString(RecordSet01.Fields.Item(2).Value);

//					//무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
//					if (P_WorkType == "A01" | P_WorkType == "A02" | Strings.Left(P_WorkType, 1) == "F" | P_WorkType == "D11") {
//						DangerCd = "";
//					} else {
//						////전문직, 계약직 이며 연봉제가 아니고 위해코드가 없으면 위해코드를 기타로..
//						if ((JIGTYP == "04" | JIGTYP == "05") & PAYTYP != "1" & JIGCOD != "73" & string.IsNullOrEmpty(DangerCd)) {
//							//창원사업장만 적용(2013.09.30 송명규 추가)
//							if (CLTCOD == "1") {
//								DangerCd = "31";
//							}
//						}
//					}

//					sQry = "            EXEC [PH_PY677_02] ";
//					sQry = sQry + "'" + CLTCOD + "',";
//					//사업장
//					sQry = sQry + "'" + PosDate + "',";
//					//일자
//					sQry = sQry + "'" + MSTCOD + "',";
//					//사번
//					sQry = sQry + "'" + ShiftDat + "',";
//					//근무형태
//					sQry = sQry + "'" + GNMUJO + "',";
//					//근무조
//					sQry = sQry + "'" + P_WorkType + "',";
//					//근태구분
//					sQry = sQry + "'" + P_Confirm + "',";
//					//기찰정상확인
//					sQry = sQry + "'" + P_GetTime + "',";
//					//출근시각
//					sQry = sQry + "'" + P_OffDt + "',";
//					//퇴근일자
//					sQry = sQry + "'" + P_OffTime + "',";
//					//퇴근시각
//					sQry = sQry + "'" + P_Base + "',";
//					//기본
//					sQry = sQry + "'" + P_Extend + "',";
//					//연장
//					sQry = sQry + "'" + P_Special + "',";
//					//특근
//					sQry = sQry + "'" + P_SpExtend + "',";
//					//특연
//					sQry = sQry + "'" + P_Midnight + "',";
//					//심야
//					sQry = sQry + "'" + P_EarlyTo + "',";
//					//조출
//					sQry = sQry + "'" + P_SEarlyTo + "',";
//					//휴조
//					sQry = sQry + "'" + P_EducTran + "',";
//					//교육훈련
//					sQry = sQry + "'" + P_LateTo + "',";
//					//지각
//					sQry = sQry + "'" + P_EarlyOff + "',";
//					//조퇴
//					sQry = sQry + "'" + P_GoOut + "',";
//					//외출
//					sQry = sQry + "'" + P_Comment + "',";
//					//비고
//					sQry = sQry + "'" + WkAbCls + "',";
//					//근태이상분류
//					sQry = sQry + "'" + WkAbCmt + "',";
//					//근태이상사유
//					sQry = sQry + "'" + ActText + "',";
//					//근무내용
//					sQry = sQry + "'" + RotateYN + "',";
//					//교대인정 황영수(2019.01.31)
//					sQry = sQry + "'" + DangerCd + "'";
//					//위해코드

//					RecordSet01.DoQuery(sQry);
//				}
//			}

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			MDC_Com.MDC_GF_Message(ref "수정 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY677_UpdateData_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			MDC_Com.MDC_GF_Message(ref "PH_PY677_UpdateData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public bool PH_PY677_AddData()
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : PH_PY677_AddData()
//			//해당모듈    : PH_PY677
//			//기능        : 데이터 INSERT
//			//인수        : 없음
//			//반환값      : 성공여부
//			//특이사항    : 이 클래스에서는 사용 안함
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;

//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			SAPbobsCOM.Recordset RecordSet02 = null;
//			RecordSet02 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			short DocEntry = 0;
//			//관리번호
//			string CLTCOD = null;
//			//사업장
//			string DestNo1 = null;
//			//출장번호1
//			string DestNo2 = null;
//			//출장번호2
//			string MSTCOD = null;
//			//사원번호
//			string MSTNAM = null;
//			//사원성명
//			string Destinat = null;
//			//출장지
//			string Dest2 = null;
//			//출장지상세
//			string CoCode = null;
//			//작번
//			string FrDate = null;
//			//시작일자
//			string FrTime = null;
//			//시작시각
//			string ToDate = null;
//			//종료일자
//			string ToTime = null;
//			//종료시각
//			//UPGRADE_NOTE: Object이(가) Object_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string Object_Renamed = null;
//			//목적
//			string Comments = null;
//			//비고
//			string RegCls = null;
//			//등록구분
//			string ObjCls = null;
//			//목적구분
//			string DestCode = null;
//			//출장지역
//			string DestDiv = null;
//			//출장구분
//			string Vehicle = null;
//			//차량구분
//			double FuelPrc = 0;
//			//1L단가
//			string FuelType = null;
//			//유류
//			double Distance = 0;
//			//거리
//			double TransExp = 0;
//			//교통비
//			double DayExp = 0;
//			//일비
//			string FoodNum = null;
//			//식수
//			double FoodExp = 0;
//			//식비
//			double ParkExp = 0;
//			//주차비
//			double TollExp = 0;
//			//도로비
//			double TotalExp = 0;
//			//합계
//			string UserSign = null;
//			//UserSign

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DestNo1 = Strings.Trim(oForm.Items.Item("DestNo1").Specific.VALUE);
//			//출장번호1
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DestNo2 = Strings.Trim(oForm.Items.Item("DestNo2").Specific.VALUE);
//			//출장번호2
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
//			//사원번호
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTNAM = Strings.Trim(oForm.Items.Item("MSTNAM").Specific.VALUE);
//			//사원성명
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Destinat = Strings.Trim(oForm.Items.Item("Destinat").Specific.VALUE);
//			//출장지
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Dest2 = Strings.Trim(oForm.Items.Item("Dest2").Specific.VALUE);
//			//출장지상세
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CoCode = Strings.Trim(oForm.Items.Item("CoCode").Specific.VALUE);
//			//작번
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FrDate = Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE);
//			//시작일자
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FrTime = Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE);
//			//시작시각
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ToDate = Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE);
//			//종료일자
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ToTime = Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE);
//			//종료시각
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Object_Renamed = Strings.Trim(oForm.Items.Item("Object").Specific.VALUE);
//			//목적
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Comments = Strings.Trim(oForm.Items.Item("Comments").Specific.VALUE);
//			//비고
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			RegCls = Strings.Trim(oForm.Items.Item("RegCls").Specific.VALUE);
//			//등록구분
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ObjCls = Strings.Trim(oForm.Items.Item("ObjCls").Specific.VALUE);
//			//목적구분
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DestCode = Strings.Trim(oForm.Items.Item("DestCode").Specific.VALUE);
//			//출장지역
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DestDiv = Strings.Trim(oForm.Items.Item("DestDiv").Specific.VALUE);
//			//출장구분
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Vehicle = Strings.Trim(oForm.Items.Item("Vehicle").Specific.VALUE);
//			//차량구분
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FuelPrc = Convert.ToDouble(Strings.Trim(oForm.Items.Item("FuelPrc").Specific.VALUE));
//			//1L단가
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FuelType = Strings.Trim(oForm.Items.Item("FuelType").Specific.VALUE);
//			//유류
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Distance = Convert.ToDouble(Strings.Trim(oForm.Items.Item("Distance").Specific.VALUE));
//			//거리
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TransExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("TransExp").Specific.VALUE));
//			//교통비
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DayExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("DayExp").Specific.VALUE));
//			//일비
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FoodNum = Strings.Trim(oForm.Items.Item("FoodNum").Specific.VALUE);
//			//식수
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FoodExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("FoodExp").Specific.VALUE));
//			//식비
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ParkExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("ParkExp").Specific.VALUE));
//			//주차비
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TollExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("TollExp").Specific.VALUE));
//			//도로비
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TotalExp = Convert.ToDouble(Strings.Trim(oForm.Items.Item("TotalExp").Specific.VALUE));
//			//합계
//			UserSign = Convert.ToString(MDC_Globals.oCompany.UserSignature);

//			//DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
//			sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY677A]";
//			RecordSet01.DoQuery(sQry);

//			if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0) {
//				DocEntry = 1;
//			} else {
//				DocEntry = Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1;
//			}

//			sQry = "            EXEC [PH_PY677_02] ";
//			sQry = sQry + "'" + DocEntry + "',";
//			//관리번호
//			sQry = sQry + "'" + CLTCOD + "',";
//			//사업장
//			sQry = sQry + "'" + DestNo1 + "',";
//			//출장번호1
//			sQry = sQry + "'" + DestNo2 + "',";
//			//출장번호2
//			sQry = sQry + "'" + MSTCOD + "',";
//			//사원번호
//			sQry = sQry + "'" + MSTNAM + "',";
//			//사원성명
//			sQry = sQry + "'" + Destinat + "',";
//			//출장지
//			sQry = sQry + "'" + Dest2 + "',";
//			//출장지상세
//			sQry = sQry + "'" + CoCode + "',";
//			//작번
//			sQry = sQry + "'" + FrDate + "',";
//			//시작일자
//			sQry = sQry + "'" + FrTime + "',";
//			//시작시각
//			sQry = sQry + "'" + ToDate + "',";
//			//종료일자
//			sQry = sQry + "'" + ToTime + "',";
//			//종료시각
//			sQry = sQry + "'" + Object_Renamed + "',";
//			//목적
//			sQry = sQry + "'" + Comments + "',";
//			//비고
//			sQry = sQry + "'" + RegCls + "',";
//			//등록구분
//			sQry = sQry + "'" + ObjCls + "',";
//			//목적구분
//			sQry = sQry + "'" + DestCode + "',";
//			//출장지역
//			sQry = sQry + "'" + DestDiv + "',";
//			//출장구분
//			sQry = sQry + "'" + Vehicle + "',";
//			//차량구분
//			sQry = sQry + "'" + FuelPrc + "',";
//			//1L단가
//			sQry = sQry + "'" + FuelType + "',";
//			//유류
//			sQry = sQry + "'" + Distance + "',";
//			//거리
//			sQry = sQry + "'" + TransExp + "',";
//			//교통비
//			sQry = sQry + "'" + DayExp + "',";
//			//일비
//			sQry = sQry + "'" + FoodNum + "',";
//			//식수
//			sQry = sQry + "'" + FoodExp + "',";
//			//식비
//			sQry = sQry + "'" + ParkExp + "',";
//			//주차비
//			sQry = sQry + "'" + TollExp + "',";
//			//도로비
//			sQry = sQry + "'" + TotalExp + "',";
//			//합계
//			sQry = sQry + "'" + UserSign + "'";
//			//UserSign

//			RecordSet02.DoQuery(sQry);

//			MDC_Com.MDC_GF_Message(ref "등록 완료!", ref "S");

//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet02 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY677_AddData_Error:

//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			functionReturnValue = false;
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet02 = null;
//			MDC_Com.MDC_GF_Message(ref "PH_PY677_AddData_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		private bool PH_PY677_HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : PH_PY677_HeaderSpaceLineDel()
//			//해당모듈    : PH_PY677
//			//기능        : 필수입력사항 체크
//			//인수        : 없음
//			//반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short ErrNum = 0;
//			ErrNum = 0;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("DestNo1").Specific.VALUE)):
//					//출장번호1
//					ErrNum = 1;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("DestNo2").Specific.VALUE)):
//					//출장번호2
//					ErrNum = 2;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE)):
//					//사원번호
//					ErrNum = 3;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE)):
//					//시작일자
//					ErrNum = 4;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE)):
//					//시작시각
//					ErrNum = 5;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE)):
//					//종료일자
//					ErrNum = 6;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE)):
//					//종료시각
//					ErrNum = 7;
//					goto PH_PY677_HeaderSpaceLineDel_Error;
//					break;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY677_HeaderSpaceLineDel_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "출장번호1은 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("DestNo1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "출장번호2는 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("DestNo2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 3) {
//				MDC_Com.MDC_GF_Message(ref "사원번호는 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 4) {
//				MDC_Com.MDC_GF_Message(ref "시작일자는 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 5) {
//				MDC_Com.MDC_GF_Message(ref "시작시각은 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 6) {
//				MDC_Com.MDC_GF_Message(ref "종료일자는 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 7) {
//				MDC_Com.MDC_GF_Message(ref "종료시각은 필수사항입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

///// 메트릭스 필수 사항 check
//		private bool PH_PY677_MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			short ErrNum = 0;
//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY677_MatrixSpaceLineDel_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 3) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 4) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 5) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY677_MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void PH_PY677_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			short ErrNum = 0;
//			string sQry = null;
//			string ItemCode = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string CLTCOD = null;
//			string TeamCode = null;
//			string RspCode = null;

//			string PreWorkType = null;
//			string WorkType = null;
//			double JanQty = 0;
//			string ymd = null;
//			string MSTCOD = null;
//			string YY = null;

//			SAPbouiCOM.ProgressBar ProgBar01 = null;

//			oForm.Freeze(true);

//			switch (oUID) {

//				case "Mat01":

//					PreWorkType = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1));

//					oMat01.FlushToDataSource();

//					if (oCol == "P_GetTime") {

//						PH_PY677_Time_ReSet(oRow);
//						oMat01.LoadFromDataSource();
//						oMat01.Columns.Item(oCol).Cells.Item(oRow).Click();

//					} else if (oCol == "P_OffTime") {

//						if (oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1) != "0000") {
//							PH_PY677_Time_Calc_Main(oDS_PH_PY677B.GetValue("U_ColTm02", oRow - 1), oRow);
//							oMat01.LoadFromDataSource();
//							oMat01.Columns.Item(oCol).Cells.Item(oRow).Click();
//						}

//					} else if (oCol == "P_WorkType") {

//						WorkType = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg21", oRow - 1));

//						switch (WorkType) {

//							case "A01":
//							case "A02":
//							case "E02":
//							case "E03":
//							case "F01":
//							case "F02":
//							case "F03":
//							case "F04":
//							case "F05":
//								//무단결근, 유계결근, 무급휴일, 휴업, 병가(휴직), 신병휴직, 정직(유결), 가사휴직, 공상휴직(F01) 추가(2017.12.07)

//								oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));
//								//퇴근일자(P_OffDt)
//								oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");
//								//출근시각(P_GetTime)
//								oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");
//								//퇴근시각(P_OffTime)

//								PH_PY677_Time_ReSet(oRow);
//								oMat01.LoadFromDataSource();
//								break;

//							case "C02":
//							case "D04":
//							case "D05":
//							case "D06":
//							case "D07":
//							case "H05":
//								//훈련, 경조휴가, 하기휴가, 특별휴가, 분만휴가, 조합활동

//								oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));
//								//퇴근일자(P_OffDt)
//								oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");
//								//출근시각(P_GetTime)
//								oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");
//								//퇴근시각(P_OffTime)
//								oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(0));
//								//교대일수

//								PH_PY677_Time_ReSet(oRow);
//								oMat01.LoadFromDataSource();
//								break;

//							case "D02":
//							case "D09":
//								//연차/반차 휴가

//								//연차/반차 휴가 잔여일수 확인
//								if (WorkType == "D02") {
//									JanQty = 1;
//								} else if (WorkType == "D09") {
//									JanQty = 0.5;
//								}

//								ymd = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1));
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//								MSTCOD = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1));

//								ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet01.RecordCount, false);

//								sQry = "      EXEC [PH_PY775_01] '";
//								sQry = sQry + CLTCOD + "','";
//								sQry = sQry + Strings.Left(ymd, 4) + "','";
//								sQry = sQry + MSTCOD + "'";

//								oRecordSet01.DoQuery(sQry);

//								if (oRecordSet01.Fields.Item("jandd").Value < JanQty) {

//									ErrNum = 1;
//									oDS_PH_PY677B.SetValue("U_ColReg21", oRow - 1, "A00");

//									oMat01.LoadFromDataSource();

//									goto PH_PY677_FlushToItemValue_Error;

//								} else {

//									oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));
//									//퇴근일자(P_OffDt)
//									oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");
//									//출근시각(P_GetTime)
//									oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");
//									//퇴근시각(P_OffTime)
//									oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(1));
//									//교대일수

//									PH_PY677_Time_ReSet(oRow);

//									oMat01.LoadFromDataSource();

//								}

//								ProgBar01.Value = 100;
//								ProgBar01.Stop();
//								break;

//							case "D08":
//							case "D10":
//								//근속보전휴가, 근속보전반차(기계사업부)

//								//근속보전휴가 잔량 확인
//								ymd = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1));
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//								MSTCOD = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg05", oRow - 1));

//								if (Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", oRow - 1)) >= Strings.Left(ymd, 4) + "0701") {
//									YY = Strings.Left(ymd, 4);
//								} else {
//									YY = Convert.ToString(Convert.ToInt16(Strings.Left(ymd, 4)) - 1);
//								}

//								ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet01.RecordCount, false);

//								sQry = "        SELECT     Bqty = ISNULL(( ";
//								sQry = sQry + "                               SELECT     ISNULL(c.Qty,0) ";
//								sQry = sQry + "                               FROM       [ZPH_GUNSOKCHA] c";
//								sQry = sQry + "                               WHERE      c.Year = '" + YY + "'";
//								sQry = sQry + "                                          AND (";
//								sQry = sQry + "                                                   CASE ";
//								sQry = sQry + "                                                       WHEN ISDATE(a.U_GrpDat) = 0 THEN a.U_StartDat ";
//								sQry = sQry + "                                                       ELSE A.U_GrpDat ";
//								sQry = sQry + "                                                   END";
//								sQry = sQry + "                                              ) BETWEEN c.DocDateFr AND c.DocDateTo";
//								sQry = sQry + "                           ),0), ";
//								//발생수량
//								sQry = sQry + "            Sqty = ISNULL(( ";
//								sQry = sQry + "                               SELECT     COUNT(*)";
//								sQry = sQry + "                               FROM       [ZPH_PY008] c";
//								sQry = sQry + "                               WHERE      c.CLTCOD = A.U_CLTCOD";
//								sQry = sQry + "                                          AND c.PosDate BETWEEN '" + YY + "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" + YY + "') + 1 ) + '0630'";
//								sQry = sQry + "                                          AND c.MSTCOD  = a.Code";
//								sQry = sQry + "                                          AND c.WorkType = 'D08'";
//								sQry = sQry + "                                          AND c.PosDate <> '" + ymd + "'";
//								sQry = sQry + "                          ),0) ";
//								//보전년차 사용수량
//								sQry = sQry + "                    + ";
//								sQry = sQry + "                    ISNULL(( ";
//								sQry = sQry + "                               SELECT     COUNT(*) / 2.0";
//								sQry = sQry + "                               FROM       [ZPH_PY008] c";
//								sQry = sQry + "                               WHERE      c.CLTCOD = A.U_CLTCOD";
//								sQry = sQry + "                                          AND c.PosDate BETWEEN '" + YY + "' + '0701' AND CONVERT(CHAR(4),CONVERT(INTEGER,'" + YY + "') + 1 ) + '0630'";
//								sQry = sQry + "                                          AND c.MSTCOD = a.Code";
//								sQry = sQry + "                                          AND c.WorkType = 'D10'";
//								sQry = sQry + "                                          AND c.PosDate <> '" + ymd + "'";
//								sQry = sQry + "                           ),0)";
//								//보전반차 사용수량
//								sQry = sQry + " FROM       [@PH_PY001A] a";
//								sQry = sQry + " WHERE      a.U_CLTCOD = '" + CLTCOD + "'";
//								sQry = sQry + "            AND a.Code = '" + MSTCOD + "'";

//								oRecordSet01.DoQuery(sQry);

//								if (WorkType == "D08") {
//									JanQty = 1;
//								} else if (WorkType == "D10") {
//									JanQty = 0.5;
//								}

//								if (oRecordSet01.Fields.Item("Bqty").Value - oRecordSet01.Fields.Item("Sqty").Value < JanQty) {

//									ErrNum = 2;
//									oDS_PH_PY677B.SetValue("U_ColReg21", oRow - 1, "A00");
//									oMat01.LoadFromDataSource();

//									goto PH_PY677_FlushToItemValue_Error;

//								} else {

//									oDS_PH_PY677B.SetValue("U_ColDt02", oRow - 1, oDS_PH_PY677B.GetValue("U_ColDt01", oRow - 1));
//									//퇴근일자(P_OffDt)
//									oDS_PH_PY677B.SetValue("U_ColTm01", oRow - 1, "00:00");
//									//출근시각(P_GetTime)
//									oDS_PH_PY677B.SetValue("U_ColTm02", oRow - 1, "00:00");
//									//퇴근시각(P_OffTime)
//									oDS_PH_PY677B.SetValue("U_ColQty12", oRow - 1, Convert.ToString(1));
//									//교대일수

//									PH_PY677_Time_ReSet(oRow);

//									oMat01.LoadFromDataSource();

//								}

//								ProgBar01.Value = 100;
//								ProgBar01.Stop();
//								break;

//						}
//					}
//					break;

//				//            Call oMat01.LoadFromDataSource
//				//            Call oMat01.AutoResizeColumns


//				case "MSTCOD":

//					//UPGRADE_WARNING: oForm.Items(MSTNAM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("MSTNAM").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'");
//					//성명
//					break;

//				case "ShiftDatCd":

//					//UPGRADE_WARNING: oForm.Items(ShiftDatNm).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("ShiftDatNm").Specific.VALUE = MDC_GetData.Get_ReData(ref "U_CodeNm", ref "U_Code", ref "[@PS_HR200L] AS T0", ref "'" + oForm.Items.Item("ShiftDatCd").Specific.VALUE + "'", ref " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'");
//					//근무형태
//					break;

//				case "GNMUJOCd":

//					//UPGRADE_WARNING: oForm.Items(GNMUJONm).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("GNMUJONm").Specific.VALUE = MDC_GetData.Get_ReData(ref "U_CodeNm", ref "U_Code", ref "[@PS_HR200L] AS T0", ref "'" + oForm.Items.Item("GNMUJOCd").Specific.VALUE + "'", ref " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'");
//					//근무조
//					break;

//				case "CLTCOD":

//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);

//					//UPGRADE_WARNING: oForm.Items(TeamCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0) {
//						//UPGRADE_WARNING: oForm.Items(TeamCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1) {
//							//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//						}
//					}

//					//부서콤보세팅
//					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
//					sQry = "        SELECT    U_Code AS [Code],";
//					sQry = sQry + "           U_CodeNm As [Name]";
//					sQry = sQry + " FROM      [@PS_HR200L]";
//					sQry = sQry + " WHERE     Code = '1'";
//					sQry = sQry + "           AND U_UseYN = 'Y'";
//					sQry = sQry + "           AND U_Char2 = '" + CLTCOD + "'";
//					sQry = sQry + " ORDER BY  U_Seq";
//					MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("TeamCode").Specific), ref sQry, ref "", ref false, ref false);
//					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//					break;

//				case "TeamCode":

//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.VALUE);

//					//UPGRADE_WARNING: oForm.Items(RspCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0) {
//						//UPGRADE_WARNING: oForm.Items(RspCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1) {
//							//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//						}
//					}

//					//담당콤보세팅
//					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
//					sQry = "        SELECT    U_Code AS [Code],";
//					sQry = sQry + "           U_CodeNm As [Name]";
//					sQry = sQry + " FROM      [@PS_HR200L]";
//					sQry = sQry + " WHERE     Code = '2'";
//					sQry = sQry + "           AND U_UseYN = 'Y'";
//					sQry = sQry + "           AND U_Char1 = '" + TeamCode + "'";
//					sQry = sQry + " ORDER BY  U_Seq";
//					MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("RspCode").Specific), ref sQry, ref "", ref false, ref false);
//					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//					break;

//				case "RspCode":

//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.VALUE);
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					RspCode = Strings.Trim(oForm.Items.Item("RspCode").Specific.VALUE);

//					//UPGRADE_WARNING: oForm.Items(ClsCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0) {
//						//UPGRADE_WARNING: oForm.Items(ClsCode).Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1) {
//							//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//						}
//					}

//					//반콤보세팅
//					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "전체");
//					sQry = "        SELECT    U_Code AS [Code],";
//					sQry = sQry + "           U_CodeNm As [Name]";
//					sQry = sQry + " FROM      [@PS_HR200L]";
//					sQry = sQry + " WHERE     Code = '9'";
//					sQry = sQry + "           AND U_UseYN = 'Y'";
//					sQry = sQry + "           AND U_Char1 = '" + RspCode + "'";
//					sQry = sQry + "           AND U_Char2 = '" + TeamCode + "'";
//					sQry = sQry + " ORDER BY  U_Seq";
//					MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("ClsCode").Specific), ref sQry, ref "", ref false, ref false);
//					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//					break;

//			}

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			return;
//			PH_PY677_FlushToItemValue_Error:

//			oForm.Freeze(false);
//			//    Call ProgBar01.Stop
//			//    ProgBar01.VALUE = 100
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("연차휴가 잔여일수가 없습니다. 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			} else if (ErrNum == 2) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("보전휴가 잔여일수가 없습니다. 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY677_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}

//		}

/////폼의 아이템 사용지정
//		public void PH_PY677_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//			}

//			return;
//			PH_PY677_FormItemEnabled_Error:

//			MDC_Com.MDC_GF_Message(ref "PH_PY677_FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

/////아이템 변경 이벤트
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;
//				////et_FORM_ACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;
//				////et_FORM_DEACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////20
//					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;
//				////et_LOST_FOCUS
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//			}
//			return;
//			Raise_FormItemEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////BeforeAction = True
//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					case "1282":
//						//추가
//						///추가버튼 클릭시 메트릭스 insertrow

//						//                Call PH_PY677_FormReset

//						//                oMat01.Clear
//						//                oMat01.FlushToDataSource
//						//                oMat01.LoadFromDataSource

//						//                oForm.Mode = fm_ADD_MODE
//						//                BubbleEvent = False
//						//                Call PH_PY677_LoadCaption

//						//oForm.Items("GCode").Click ct_Regular


//						return;

//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
//						PH_PY677_Add_MatrixRow(oMat01.VisualRowCount);
//						break;

//				}
//			////BeforeAction = False
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					////Call PH_PY677_FormItemEnabled '//UDO방식
//					case "1282":
//						//추가
//						break;
//					//                oMat01.Clear
//					//                oDS_PH_PY677A.Clear

//					//                Call PH_PY677_LoadCaption
//					//                Call PH_PY677_FormItemEnabled
//					////Call PH_PY677_FormItemEnabled '//UDO방식
//					////Call PH_PY677_AddMatrixRow(0, True) '//UDO방식
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;
//					////Call PH_PY677_FormItemEnabled

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 이후 처리
//						oForm.Freeze(true);
//						oDS_PH_PY677B.RemoveRecord(oDS_PH_PY677B.Size - 1);
//						oMat01.LoadFromDataSource();
//						oForm.Freeze(false);
//						break;

//				}
//			}
//			return;
//			Raise_FormMenuEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////BeforeAction = True
//			if ((BusinessObjectInfo.BeforeAction == true)) {
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
//			////BeforeAction = False
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			return;
//			Raise_FormDataEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_RightClickEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				if (pval.ItemUID == "PH_PY677") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}

//				//수정 버튼
//				if (pval.ItemUID == "BtnUpdate") {

//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

//						PH_PY677_UpdateData();

//						//                If PH_PY677_HeaderSpaceLineDel() = False Then
//						//                    BubbleEvent = False
//						//                    Exit Sub
//						//                End If
//						//
//						//'                If PH_PY677_DataCheck() = False Then
//						//'                    BubbleEvent = False
//						//'                    Exit Sub
//						//'                End If
//						//
//						//'                If PH_PY677_AddData() = False Then
//						//'                    BubbleEvent = False
//						//'                    Exit Sub
//						//'                End If
//						//
//						//'                Call PH_PY677_FormReset
//						//                oForm.Mode = fm_ADD_MODE
//						//
//						//'                Call PH_PY677_LoadCaption
//						//                Call PH_PY677_MTX01
//						//
//						//                oLast_Mode = oForm.Mode

//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {

//						if (PH_PY677_HeaderSpaceLineDel() == false) {
//							BubbleEvent = false;
//							return;
//						}

//						//                If PH_PY677_DataCheck() = False Then
//						//                    BubbleEvent = False
//						//                    Exit Sub
//						//                End If

//						//                If PH_PY677_UpdateData() = False Then
//						//                    BubbleEvent = False
//						//                    Exit Sub
//						//                End If
//						//
//						//                Call PH_PY677_FormReset
//						//                oForm.Mode = fm_ADD_MODE
//						//
//						//                Call PH_PY677_LoadCaption
//						//                Call PH_PY677_MTX01

//						//                oForm.Items("GCode").Click ct_Regular
//					}

//				///조회
//				} else if (pval.ItemUID == "BtnSearch") {

//					//            Call PH_PY677_FormReset
//					//            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE

//					//            Call PH_PY677_LoadCaption
//					PH_PY677_MTX01();

//					//        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
//					//
//					//            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
//					//
//					//                Call PH_PY677_DeleteData
//					//                Call PH_PY677_FormReset
//					//                oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
//					//
//					//                Call PH_PY677_LoadCaption
//					//                Call PH_PY677_MTX01
//					//
//					//            Else
//					//
//					//            End If

//					//        ElseIf pval.ItemUID = "BtnPrint" Then
//					//
//					//            Call PH_PY677_Print_Report01

//				} else if (pval.ItemUID == "BtnChk") {

//					//            If oForm.Items("BtnChk").Specific.Caption = "전체선택" Then
//					//                oForm.Items("BtnChk").Specific.Caption = "전체해제"
//					//            Else
//					//                oForm.Items("BtnChk").Specific.Caption = "전체선택"
//					//            End If

//					PH_PY677_CheckAll();

//				} else if (pval.ItemUID == "BtnConfirm") {

//					PH_PY677_ConfirmAll();

//				}

//			} else if (pval.BeforeAction == false) {
//				if (pval.ItemUID == "PH_PY677") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}
//			}

//			return;
//			Raise_EVENT_ITEM_PRESSED_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "MSTCOD", "");
//				//사번
//				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "ShiftDatCd", "");
//				//근무형태
//				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "GNMUJOCd", "");
//				//근무조

//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_KEY_DOWN_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				if (pval.ItemUID == "Mat01") {

//					if (pval.Row > 0) {

//						oMat01.SelectRow(pval.Row, true, false);

//						//                Call oForm.Freeze(True)
//						//
//						//                'DataSource를 이용하여 각 컨트롤에 값을 출력
//						//                Call oDS_PH_PY677A.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
//						//                Call oDS_PH_PY677A.setValue("U_CLTCOD", 0, oMat01.Columns("CLTCOD").Cells(pval.Row).Specific.VALUE) '사업장
//						//                Call oDS_PH_PY677A.setValue("U_DestNo1", 0, oMat01.Columns("DestNo1").Cells(pval.Row).Specific.VALUE) '출장번호1
//						//                Call oDS_PH_PY677A.setValue("U_DestNo2", 0, oMat01.Columns("DestNo2").Cells(pval.Row).Specific.VALUE) '출장번호2
//						//                Call oDS_PH_PY677A.setValue("U_MSTCOD", 0, oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE) '사원번호
//						//                Call oDS_PH_PY677A.setValue("U_MSTNAM", 0, oMat01.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE) '사원성명
//						//                Call oDS_PH_PY677A.setValue("U_Destinat", 0, oMat01.Columns("Destinat").Cells(pval.Row).Specific.VALUE) '출장지
//						//                Call oDS_PH_PY677A.setValue("U_Dest2", 0, oMat01.Columns("Dest2").Cells(pval.Row).Specific.VALUE) '출장지상세
//						//                Call oDS_PH_PY677A.setValue("U_CoCode", 0, oMat01.Columns("CoCode").Cells(pval.Row).Specific.VALUE) '작번
//						//                Call oDS_PH_PY677A.setValue("U_FrDate", 0, Replace(oMat01.Columns("FrDate").Cells(pval.Row).Specific.VALUE, ".", "")) '시작일자
//						//                Call oDS_PH_PY677A.setValue("U_FrTime", 0, oMat01.Columns("FrTime").Cells(pval.Row).Specific.VALUE) '시작시각
//						//                Call oDS_PH_PY677A.setValue("U_ToDate", 0, Replace(oMat01.Columns("ToDate").Cells(pval.Row).Specific.VALUE, ".", "")) '종료일자
//						//                Call oDS_PH_PY677A.setValue("U_ToTime", 0, oMat01.Columns("ToTime").Cells(pval.Row).Specific.VALUE) '종료시각
//						//                Call oDS_PH_PY677A.setValue("U_Object", 0, oMat01.Columns("Object").Cells(pval.Row).Specific.VALUE) '목적
//						//                Call oDS_PH_PY677A.setValue("U_Comments", 0, oMat01.Columns("Comments").Cells(pval.Row).Specific.VALUE) '비고
//						//                Call oDS_PH_PY677A.setValue("U_RegCls", 0, oMat01.Columns("RegCls").Cells(pval.Row).Specific.VALUE) '등록구분
//						//                Call oDS_PH_PY677A.setValue("U_ObjCls", 0, oMat01.Columns("ObjCls").Cells(pval.Row).Specific.VALUE) '목적구분
//						//                Call oDS_PH_PY677A.setValue("U_DestCode", 0, oMat01.Columns("DestCode").Cells(pval.Row).Specific.VALUE) '출장지역
//						//                Call oDS_PH_PY677A.setValue("U_DestDiv", 0, oMat01.Columns("DestDiv").Cells(pval.Row).Specific.VALUE) '출장구분
//						//                Call oDS_PH_PY677A.setValue("U_Vehicle", 0, oMat01.Columns("Vehicle").Cells(pval.Row).Specific.VALUE) '차량구분
//						//                Call oDS_PH_PY677A.setValue("U_FuelPrc", 0, oMat01.Columns("FuelPrc").Cells(pval.Row).Specific.VALUE) '1L단가
//						//                Call oDS_PH_PY677A.setValue("U_FuelType", 0, oMat01.Columns("FuelType").Cells(pval.Row).Specific.VALUE) '유류
//						//                Call oDS_PH_PY677A.setValue("U_Distance", 0, oMat01.Columns("Distance").Cells(pval.Row).Specific.VALUE) '거리
//						//                Call oDS_PH_PY677A.setValue("U_TransExp", 0, oMat01.Columns("TransExp").Cells(pval.Row).Specific.VALUE) '교통비
//						//                Call oDS_PH_PY677A.setValue("U_DayExp", 0, oMat01.Columns("DayExp").Cells(pval.Row).Specific.VALUE) '일비
//						//                Call oDS_PH_PY677A.setValue("U_FoodNum", 0, oMat01.Columns("FoodNum").Cells(pval.Row).Specific.VALUE) '식수
//						//                Call oDS_PH_PY677A.setValue("U_FoodExp", 0, oMat01.Columns("FoodExp").Cells(pval.Row).Specific.VALUE) '식비
//						//                Call oDS_PH_PY677A.setValue("U_ParkExp", 0, oMat01.Columns("ParkExp").Cells(pval.Row).Specific.VALUE) '주차비
//						//                Call oDS_PH_PY677A.setValue("U_TollExp", 0, oMat01.Columns("TollExp").Cells(pval.Row).Specific.VALUE) '도로비
//						//                Call oDS_PH_PY677A.setValue("U_TotalExp", 0, oMat01.Columns("TotalExp").Cells(pval.Row).Specific.VALUE) '합계
//						//
//						//                oForm.Mode = fm_UPDATE_MODE
//						//                Call PH_PY677_LoadCaption
//						//
//						//                Call oForm.Freeze(False)

//					}
//				}
//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_CLICK_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//				if (pval.ItemUID == "Mat01") {

//					//            If pval.ColUID = "P_Confirm" Or pval.ColUID = "P_WorkType" Then
//					if (pval.ColUID == "P_WorkType") {

//						PH_PY677_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);

//					}

//				} else {

//					PH_PY677_FlushToItemValue(pval.ItemUID);

//				}

//			}

//			return;
//			Raise_EVENT_COMBO_SELECT_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_DOUBLE_CLICK_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			if (pval.BeforeAction == true) {

//				if (pval.ItemChanged == true) {

//					if ((pval.ItemUID == "Mat01")) {


//						if (pval.ColUID == "P_GetTime" | pval.ColUID == "P_OffTime") {

//							PH_PY677_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);

//						}

//						//                oMat01.LoadFromDataSource
//						oMat01.AutoResizeColumns();

//					} else {

//						PH_PY677_FlushToItemValue(pval.ItemUID);

//					}

//					oForm.Update();
//				}

//			} else if (pval.BeforeAction == false) {

//			}

//			oForm.Freeze(false);

//			return;
//			Raise_EVENT_VALIDATE_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY677_FormItemEnabled();
//				////Call PH_PY677_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
//			}
//			return;
//			Raise_EVENT_MATRIX_LOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY677_FormResize();
//				oMat01.AutoResizeColumns();
//			}
//			return;
//			Raise_EVENT_RESIZE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				//        If (pval.ItemUID = "ItemCode") Then
//				//            Dim oDataTable01 As SAPbouiCOM.DataTable
//				//            Set oDataTable01 = pval.SelectedObjects
//				//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
//				//            Set oDataTable01 = Nothing
//				//        End If
//				//        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
//				//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY677A", "U_CardCode,U_CardName")
//				//        End If
//			}
//			return;
//			Raise_EVENT_CHOOSE_FROM_LIST_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_EVENT_GOT_FOCUS_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//				SubMain.RemoveForms(oFormUniqueID01);
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oMat01 = null;
//			}
//			return;
//			Raise_EVENT_FORM_UNLOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			if ((oLastColRow01 > 0)) {
//				if (pval.BeforeAction == true) {
//					//            If (PH_PY677_Validate("행삭제") = False) Then
//					//                BubbleEvent = False
//					//                Exit Sub
//					//            End If
//					////행삭제전 행삭제가능여부검사
//				} else if (pval.BeforeAction == false) {
//					for (i = 1; i <= oMat01.VisualRowCount; i++) {
//						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
//					}
//					oMat01.FlushToDataSource();
//					oDS_PH_PY677A.RemoveRecord(oDS_PH_PY677A.Size - 1);
//					oMat01.LoadFromDataSource();
//					if (oMat01.RowCount == 0) {
//						PH_PY677_Add_MatrixRow(0);
//					} else {
//						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY677A.GetValue("U_CntcCode", oMat01.RowCount - 1)))) {
//							PH_PY677_Add_MatrixRow(oMat01.RowCount);
//						}
//					}
//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY677_CreateItems()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			string oQuery01 = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    Set oDS_PH_PY677A = oForm.DataSources.DBDataSources("@PH_PY677A")
//			oDS_PH_PY677B = oForm.DataSources.DBDataSources("@PS_USERDS01");

//			//// 메트릭스 개체 할당
//			oMat01 = oForm.Items.Item("Mat01").Specific;
//			oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat01.AutoResizeColumns();

//			//사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

//			//시작일자
//			oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");

//			//종료일자
//			oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");

//			//부서
//			oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

//			//담당
//			oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

//			//반
//			oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

//			//근무형태 (코드)
//			oForm.DataSources.UserDataSources.Add("ShiftDatCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ShiftDatCd").Specific.DataBind.SetBound(true, "", "ShiftDatCd");

//			//근무형태(명)
//			oForm.DataSources.UserDataSources.Add("ShiftDatNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ShiftDatNm").Specific.DataBind.SetBound(true, "", "ShiftDatNm");

//			//근무조(코드)
//			oForm.DataSources.UserDataSources.Add("GNMUJOCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("GNMUJOCd").Specific.DataBind.SetBound(true, "", "GNMUJOCd");

//			//근무조(명)
//			oForm.DataSources.UserDataSources.Add("GNMUJONm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("GNMUJONm").Specific.DataBind.SetBound(true, "", "GNMUJONm");

//			//사원번호
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

//			//사원성명
//			oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

//			//근태기찰정상확인
//			oForm.DataSources.UserDataSources.Add("Confirm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Confirm").Specific.DataBind.SetBound(true, "", "Confirm");

//			//기찰이상구분
//			oForm.DataSources.UserDataSources.Add("Class", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Class").Specific.DataBind.SetBound(true, "", "Class");

//			//근태구분(2014.05.13 송명규 추가)
//			oForm.DataSources.UserDataSources.Add("WorkType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("WorkType").Specific.DataBind.SetBound(true, "", "WorkType");

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY677_CreateItems_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

/////콤보박스 set
//		public void PH_PY677_ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			//헤더
//			//기찰이상구분
//			oCombo = oForm.Items.Item("Class").Specific;
//			oCombo.ValidValues.Add("%", "전체");
//			oCombo.ValidValues.Add("Y", "근태이상");
//			oCombo.ValidValues.Add("N", "정상");
//			oCombo.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//근태이상자 확인
//			oCombo = oForm.Items.Item("Confirm").Specific;
//			oCombo.ValidValues.Add("%", "전체");
//			oCombo.ValidValues.Add("N", "미확인[N]");
//			oCombo.ValidValues.Add("Y", "확인[Y]");
//			oForm.Items.Item("Confirm").DisplayDesc = true;
//			oCombo.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//근태구분
//			//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("WorkType").Specific.ValidValues.Add("%", "전체");
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P221'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_SetMod.Set_ComboList(ref (oForm.Items.Item("WorkType").Specific), ref sQry, ref "", ref false, ref false);
//			//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("WorkType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			////////////매트릭스//////////

//			//근무형태
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P154'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("ShiftDat"), sQry);

//			//근무조
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P155'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("GNMUJO"), sQry);

//			//요일구분
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P202'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("DayType"), sQry);

//			//근태구분
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P221'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("P_WorkType"), sQry);

//			//기찰정상확인
//			//    Set oCombo = oMat01.Columns("P_Confirm")
//			oMat01.Columns.Item("P_Confirm").ValidValues.Add("N", "미확인[N]");
//			oMat01.Columns.Item("P_Confirm").ValidValues.Add("Y", "확인[Y]");
//			oMat01.Columns.Item("P_Confirm").DisplayDesc = true;
//			//    Call oCombo.Select(0, psk_Index)

//			//근태이상분류
//			sQry = "        SELECT    U_Code AS [Code],";
//			sQry = sQry + "           U_CodeNm As [Name]";
//			sQry = sQry + " FROM      [@PS_HR200L]";
//			sQry = sQry + " WHERE     Code = 'P237'";
//			sQry = sQry + "           AND U_UseYN = 'Y'";
//			sQry = sQry + " ORDER BY  U_Seq";
//			MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns.Item("WkAbCls"), sQry);

//			//교대인정(2016.12.07
//			oMat01.Columns.Item("RotateYN").ValidValues.Add("N", "N");
//			oMat01.Columns.Item("RotateYN").ValidValues.Add("Y", "Y");

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			return;
//			PH_PY677_ComboBox_Setting_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY677_CF_ChooseFromList()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY677_CF_ChooseFromList_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_EnableMenus()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY677_EnableMenus_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY677_FormItemEnabled();
//				////Call PH_PY677_AddMatrixRow(0, True) '//UDO방식일때
//			} else {
//				//        oForm.Mode = fm_FIND_MODE
//				//        Call PH_PY677_FormItemEnabled
//				//        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
//				//        oForm.Items("1").Click ct_Regular
//			}
//			return;
//			PH_PY677_SetDocument_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_FormResize()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oMat01.AutoResizeColumns();

//			return;
//			PH_PY677_FormResize_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_Time_ReSet(short pRow)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			////시간 초기화
//			//    Call oForm.Freeze(True)

//			oDS_PH_PY677B.SetValue("U_ColQty01", pRow - 1, Convert.ToString(0));
//			//기본
//			oDS_PH_PY677B.SetValue("U_ColQty02", pRow - 1, Convert.ToString(0));
//			//연장
//			oDS_PH_PY677B.SetValue("U_ColQty03", pRow - 1, Convert.ToString(0));
//			//심야
//			oDS_PH_PY677B.SetValue("U_ColQty06", pRow - 1, Convert.ToString(0));
//			//조출
//			oDS_PH_PY677B.SetValue("U_ColQty04", pRow - 1, Convert.ToString(0));
//			//특근
//			oDS_PH_PY677B.SetValue("U_ColQty05", pRow - 1, Convert.ToString(0));
//			//특연
//			oDS_PH_PY677B.SetValue("U_ColQty08", pRow - 1, Convert.ToString(0));
//			//교육훈련
//			oDS_PH_PY677B.SetValue("U_ColQty07", pRow - 1, Convert.ToString(0));
//			//휴조
//			oDS_PH_PY677B.SetValue("U_ColQty09", pRow - 1, Convert.ToString(0));
//			//지각
//			oDS_PH_PY677B.SetValue("U_ColQty10", pRow - 1, Convert.ToString(0));
//			//조퇴
//			oDS_PH_PY677B.SetValue("U_ColQty11", pRow - 1, Convert.ToString(0));
//			//외출

//			oDS_PH_PY677B.SetValue("U_ColTm02", pRow - 1, "00:00");
//			//퇴근시각

//			//    oForm.DataSources.UserDataSources.Item("Base").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("Extend").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("Midnight").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("EarlyTo").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("Special").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("SpExtend").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("EducTran").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("SEarlyTo").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("LateTo").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("EarlyOff").VALUE = 0
//			//    oForm.DataSources.UserDataSources.Item("GoOut").VALUE = 0
//			//
//			//    oForm.DataSources.UserDataSources.Item("GOFrTime").VALUE = "0000"
//			//    oForm.DataSources.UserDataSources.Item("GOToTime").VALUE = "0000"
//			//    oForm.DataSources.UserDataSources.Item("GOFrTim2").VALUE = "0000"
//			//    oForm.DataSources.UserDataSources.Item("GOToTim2").VALUE = "0000"

//			//    Call oForm.Freeze(False)
//			return;
//			PH_PY677_Time_ReSet_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_Time_ReSet_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private object PH_PY677_Time_Calc_Main(string OffTime, short pRow)
//		{
//			object functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			////근태구분 Select
//			short i = 0;
//			string CLTCOD = null;
//			string ShiftDat = null;
//			////근무형태
//			string GNMUJO = null;
//			////근무조
//			string DayOff = null;
//			////평일휴일구분

//			string PosDate = null;
//			////기준일
//			string GetDate = null;
//			////출근일
//			string OffDate = null;
//			////퇴근일
//			string GetTime = null;
//			////출근시간
//			//Dim OffTime As String '//퇴근시간

//			string FromTime = null;
//			string ToTime = null;

//			double hTime1 = 0;
//			////오전10분휴식시간
//			double hTime2 = 0;
//			////점심휴식시간
//			double hTime3 = 0;
//			////오후 10분 휴식시간
//			double hTime4 = 0;
//			////저녁휴식시간
//			double hTime5 = 0;
//			////야간휴식시간


//			string NextDay = null;
//			string TimeType = null;
//			string sQry = null;

//			string STime = null;
//			string ETime = null;

//			double EarlyTo = 0;
//			double SEarlyTo = 0;
//			double Base = 0;
//			double Special = 0;
//			double Extend = 0;
//			double SpExtend = 0;
//			double Midnight = 0;

//			string WorkType = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			ShiftDat = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg07", pRow - 1));
//			//근무형태
//			GNMUJO = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg08", pRow - 1));
//			//근무조

//			PosDate = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt05", pRow - 1));
//			//일자
//			GetDate = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt01", pRow - 1));
//			//출근일자
//			OffDate = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColDt02", pRow - 1));
//			//퇴근일자
//			DayOff = Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg10", pRow - 1));
//			//휴일 평일 구분
//			GetTime = Strings.Trim(Convert.ToString(oDS_PH_PY677B.GetValue("U_ColTm01", pRow - 1)));
//			//출근시각
//			OffTime = Strings.Trim(Convert.ToString(oDS_PH_PY677B.GetValue("U_ColTm02", pRow - 1)));
//			//퇴근시각

//			GetTime = "0" + GetTime;
//			GetTime = Strings.Right(GetTime, 4);

//			OffTime = "0" + OffTime;
//			OffTime = Strings.Right(OffTime, 4);

//			WorkType = Strings.Trim(Convert.ToString(oDS_PH_PY677B.GetValue("U_ColReg21", pRow - 1)));
//			//근태구분

//			sQry = "        SELECT   U_TimeType, ";
//			sQry = sQry + "          U_FromTime, ";
//			sQry = sQry + "          U_ToTime, ";
//			sQry = sQry + "          U_NextDay ";
//			sQry = sQry + " FROM     [@PH_PY002A] a ";
//			sQry = sQry + "          INNER JOIN ";
//			sQry = sQry + "          [@PH_PY002B] b ";
//			sQry = sQry + "             On a.Code = b.Code ";
//			sQry = sQry + " WHERE    a.U_CLTCOD = '" + CLTCOD + "'";
//			////사업부
//			sQry = sQry + "          AND a.U_SType = '" + ShiftDat + "'";
//			////교대
//			sQry = sQry + "          AND a.U_Shift = '" + GNMUJO + "'";
//			////조
//			sQry = sQry + "          AND b.U_DayType = '" + DayOff + "'";
//			////평일

//			oRecordSet.DoQuery(sQry);

//			if ((oRecordSet.RecordCount == 0)) {
//				//        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
//				goto PH_PY677_Time_Calc_Main_Exit;
//			}

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				FromTime = Convert.ToString(oRecordSet.Fields.Item(1).Value);
//				FromTime = "0000" + FromTime;
//				FromTime = Strings.Right(FromTime, 4);

//				ToTime = Convert.ToString(oRecordSet.Fields.Item(2).Value);
//				ToTime = "0000" + ToTime;
//				ToTime = Strings.Right(ToTime, 4);

//				NextDay = Strings.Trim(oRecordSet.Fields.Item(3).Value);
//				TimeType = Strings.Trim(oRecordSet.Fields.Item(0).Value);

//				if (string.IsNullOrEmpty(NextDay)) {
//					NextDay = "N";
//				}

//				if (NextDay == "N") {
//					if (ToTime == "0000") {
//						ToTime = "2400";
//					}
//				}

//				switch (TimeType) {
//					case "40":
//						////조출
//						////출근당일이면
//						if (GetDate == PosDate) {
//							////출근시간 < 기준종료시간
//							if (GetTime < ToTime) {
//								STime = GetTime;
//								ETime = ToTime;

//								if (DayOff == "1") {
//									EarlyTo = PH_PY677_Time_Calc(STime, ETime);
//								} else {
//									SEarlyTo = PH_PY677_Time_Calc(STime, ETime);
//								}

//							}
//						}
//						break;
//					case "10":
//					case "50":
//						////정상근무시간
//						if ((GNMUJO == "11" | GNMUJO == "21")) {
//							switch (NextDay) {
//								case "N":
//									////당일
//									////기준일자 = 출근일자
//									if (PosDate == GetDate) {
//										////1교대1조, 2교대 1조당일
//										////출근시간 < 시작시간
//										if (GetTime < FromTime) {
//											STime = FromTime;
//											////시작시간
//										} else {
//											STime = GetTime;
//											////출근시간
//										}

//										////퇴근일이 틀릴때(다음날 퇴근일때)
//										if (GetDate != OffDate) {
//											ETime = ToTime;
//											////종료시간
//										} else {
//											////퇴근시간 < 종료시간
//											if (OffTime < ToTime) {
//												ETime = OffTime;
//												////퇴근시간
//											} else {
//												ETime = ToTime;
//												////종료시간
//											}
//										}

//										if (Convert.ToDouble(DayOff) == 1) {
//											Base = Base + PH_PY677_Time_Calc(STime, ETime);
//											////평일
//										} else {
//											Special = Special + PH_PY677_Time_Calc(STime, ETime);
//											////휴일
//										}
//									}
//									break;
//								case "Y":
//									////익일
//									break;
//								//                            If PosDate = OffDate Then '//기준일자와 퇴근일자가 같으면 계산안함(당일퇴근)
//								//                                STime = "0000"
//								//                                ETime = "0000"
//								//                            Else                      '//익일퇴근일때
//								//                                If OffTime < ToTime Then  '//퇴근시간 < 종료시간
//								//                                    STime = "0000"        '//00시 부터
//								//                                    ETime = OffTime       '//퇴근시간
//								//                                Else
//								//                                    STime = "0000"
//								//                                    ETime = ToTime
//								//                                End If
//								//                            End If
//								//
//								//                            If DayOff = 1 Then
//								//                                Base = Base + PH_PY677_Time_Calc(STime, ETime)
//								//                            Else
//								//                                Special = Special + PH_PY677_Time_Calc(STime, ETime)
//								//                            End If
//							}
//						} else {
//							////당일 익일 더해야 한다.
//							if (GNMUJO == "22") {
//								switch (NextDay) {
//									case "N":
//										////당일
//										////기준일 = 출근일
//										if (PosDate == GetDate) {
//											////출근시간 < 시작시간
//											if (GetTime < FromTime) {
//												STime = FromTime;
//												////시작시간
//											} else {
//												STime = GetTime;
//												////출근시간
//											}

//											////퇴근일이 같으면(기준일 퇴근)
//											if (GetDate == OffDate) {
//												////퇴근시간 < 24시
//												if (OffTime < "2400") {
//													ETime = OffTime;
//													////퇴근시간
//												} else {
//													ETime = "2400";
//												}
//											} else {
//												ETime = "2400";
//											}

//											if (Convert.ToDouble(DayOff) == 1) {
//												Base = Base + PH_PY677_Time_Calc(STime, ETime);
//											} else {
//												Special = Special + PH_PY677_Time_Calc(STime, ETime);
//											}
//										}
//										break;

//									case "Y":
//										////익일
//										////기준일이 같으면 계산안함
//										if (PosDate == OffDate) {
//											STime = "0000";
//											ETime = "0000";
//										} else {
//											STime = "0000";
//											////시작시간은 00시

//											////퇴근시간 < 종료시간
//											if (OffTime < ToTime) {
//												ETime = OffTime;
//												////퇴근시간
//											} else {
//												ETime = ToTime;
//												////종료시간
//											}

//										}

//										if (Convert.ToDouble(DayOff) == 1) {
//											Base = Base + PH_PY677_Time_Calc(STime, ETime);
//										} else {
//											Special = Special + PH_PY677_Time_Calc(STime, ETime);
//										}
//										break;

//								}

//							}

//						}
//						break;
//					case "65":
//					case "66":
//					case "15":
//						////오전, 오후 휴식시간, 점심시간
//						////1교대1조 2교대 1조
//						if ((GNMUJO == "11" | GNMUJO == "21" | GNMUJO == "22")) {
//							switch (NextDay) {
//								case "N":
//									////당일
//									////당일 출근이 아니면
//									if (PosDate != GetDate) {
//										STime = "0000";
//										ETime = "0000";
//									} else {
//										////당일퇴근
//										if (PosDate == OffDate) {
//											////출근시간 < 시작시간
//											if (GetTime < FromTime) {
//												////시작시간 <= 퇴근시간
//												if (FromTime <= OffTime) {
//													STime = FromTime;
//													////시작시간
//													////종료시간 < 퇴근시간
//													if (ToTime < OffTime) {
//														ETime = ToTime;
//													} else {
//														ETime = OffTime;
//													}
//												} else {
//													STime = "0000";
//													ETime = "0000";
//												}
//											} else {
//												////출근시간 > 종료시간
//												if (GetTime > ToTime) {
//													STime = "0000";
//													ETime = "0000";
//												} else {
//													STime = GetTime;
//													////출근시간
//													////종료시간 < 퇴근시간
//													if (ToTime < OffTime) {
//														ETime = ToTime;
//														////종료시간
//													} else {
//														ETime = OffTime;
//														////퇴근시간
//													}
//												}
//											}
//										////다음날 퇴근
//										} else {
//											////종료시간 < 출근시간
//											if (ToTime < GetTime) {
//												STime = "0000";
//												ETime = "0000";
//											} else {
//												////출근시간 < 시작시간
//												if (GetTime < FromTime) {
//													STime = FromTime;
//													ETime = ToTime;

//												////시작시간이후 출근
//												} else {
//													STime = GetTime;
//												}
//												ETime = ToTime;
//											}

//										}
//									}
//									break;
//								case "Y":
//									////익일
//									////기준일 = 퇴근일(당일퇴근)
//									if (PosDate == OffDate) {
//										STime = "0000";
//										ETime = "0000";
//									} else {
//										////퇴근시간 < 시작시간
//										if (OffTime < FromTime) {
//											STime = "0000";
//											ETime = "0000";
//										} else {
//											STime = FromTime;
//											////퇴근시간 < 종료시간
//											if (OffTime < ToTime) {
//												ETime = OffTime;
//											} else {
//												ETime = ToTime;
//											}
//										}
//									}
//									break;
//							}

//						}

//						hTime1 = PH_PY677_Time_Calc(STime, ETime);
//						////오전휴식시간

//						////야간조(2교대) 오전휴식일경우
//						if (GNMUJO == "22" & TimeType == "65") {
//							Midnight = Midnight - hTime1;
//						} else {
//							if (DayOff == "1") {
//								Base = Base - hTime1;
//							} else {
//								Special = Special - hTime1;
//							}
//						}
//						break;

//					case "20":
//					case "60":
//						////연장근무
//						////1교대1조 2교대 1조
//						if ((GNMUJO == "11" | GNMUJO == "21")) {
//							switch (NextDay) {
//								case "N":
//									////당일
//									//// 출근일 <> 퇴근일
//									if (PosDate != OffDate) {
//										////출근시간 < 시작시간
//										if (GetTime < FromTime) {
//											STime = FromTime;
//											////
//										} else {
//											STime = GetTime;
//										}
//										ETime = "2400";
//									////당일 퇴근
//									} else {
//										////퇴근시간 < 시작시간
//										if (OffTime < FromTime) {
//											STime = "0000";
//											ETime = "0000";
//										} else {
//											STime = FromTime;
//											////종료시간
//											ETime = OffTime;
//											////퇴근시간
//										}
//									}
//									if (DayOff == "1") {
//										Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
//									} else {
//										SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
//									}
//									break;
//								case "Y":
//									////익일
//									////기준일 <> 퇴근일
//									if (PosDate != OffDate) {
//										STime = "0000";
//										////퇴근시간 < 종료시간
//										if (OffTime < ToTime) {
//											ETime = OffTime;
//										} else {
//											ETime = ToTime;
//										}

//									} else {
//										STime = "0000";
//										ETime = "0000";
//									}
//									if (DayOff == "1") {
//										Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
//									} else {
//										SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
//									}
//									break;

//							}
//						////2교대 2조
//						} else if (GNMUJO == "22") {
//							switch (NextDay) {
//								case "N":
//									////당일
//									break;
//								////연장근무없다
//								case "Y":
//									////익일
//									//// 기준일 <> 퇴근일
//									if (PosDate != OffDate) {
//										////퇴근시간 < 시작시간
//										if (OffTime < FromTime) {
//											STime = "0000";
//											ETime = "0000";
//										} else {
//											STime = FromTime;
//											ETime = OffTime;

//										}
//									} else {
//										STime = "0000";
//										ETime = "0000";
//									}

//									if (DayOff == "1") {
//										Extend = Extend + PH_PY677_Time_Calc(STime, ETime);
//									} else {
//										SpExtend = SpExtend + PH_PY677_Time_Calc(STime, ETime);
//									}
//									break;
//							}
//						}
//						break;
//					case "25":
//						////저녁휴식
//						break;
//					case "30":
//						////심야시간
//						////1교대1조 2교대 1조
//						if ((GNMUJO == "11" | GNMUJO == "21" | GNMUJO == "22")) {
//							switch (NextDay) {
//								case "N":
//									////당일
//									////기준일자 <> 퇴근일자
//									if (PosDate != OffDate) {
//										////출근시간 < 시작시간
//										if (GetTime < FromTime) {
//											STime = FromTime;
//											////시작시간
//										} else {
//											STime = GetTime;
//											////출근시간
//										}
//										ETime = "2400";
//									////당일퇴근시
//									} else {
//										////퇴근시간 < 시작시간
//										if (OffTime < FromTime) {
//											STime = "0000";
//											ETime = "0000";
//										} else {
//											////출근시간 < 시작시간
//											if (GetTime < FromTime) {
//												STime = FromTime;
//												////시작시간
//											} else {
//												STime = GetTime;
//											}

//											////퇴근시간 < 종료시간
//											if (OffTime < ToTime) {
//												ETime = OffTime;
//											} else {
//												ETime = ToTime;
//												////종료시간
//											}

//										}
//									}
//									Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime);
//									break;
//								case "Y":
//									////익일
//									////출근일 <> 퇴근일
//									if (PosDate != OffDate) {
//										STime = "0000";
//										////퇴근시간 < 종료시간
//										if (OffTime < ToTime) {
//											ETime = OffTime;
//											////퇴근시간
//										} else {
//											ETime = ToTime;
//											////종료시간
//										}
//									} else {
//										STime = "0000";
//										ETime = "0000";
//									}
//									Midnight = Midnight + PH_PY677_Time_Calc(STime, ETime);
//									break;
//							}

//						}
//						break;
//					case "35":
//						////야간휴식 수정해야함
//						////hTime5
//						switch (NextDay) {
//							case "N":
//								////당일
//								////기준일자 = 퇴근일자
//								if (PosDate == OffDate) {
//									////퇴근시간 < 시작시간
//									if (OffTime < FromTime) {
//										STime = "0000";
//										ETime = "0000";
//									} else {
//										if (GetTime < FromTime) {
//											STime = FromTime;
//										} else {
//											STime = GetTime;
//										}

//										if (OffTime < ToTime) {
//											ETime = OffTime;
//										} else {
//											ETime = ToTime;
//										}
//									}
//								} else {
//									////다음날 퇴근
//									////출근시간 < 시작시간
//									if (GetTime < FromTime) {
//										STime = FromTime;
//									} else {
//										STime = GetTime;
//									}

//									ETime = ToTime;
//								}
//								break;

//							case "Y":
//								////기준일자 = 퇴근일자
//								if (PosDate == OffDate) {
//									STime = "0000";
//									ETime = "0000";
//								} else {
//									////퇴근시간 < 시작시간
//									if (OffTime < FromTime) {
//										STime = "0000";
//										ETime = "0000";
//									} else {
//										STime = FromTime;
//										if (OffTime < ToTime) {
//											ETime = OffTime;
//										} else {
//											ETime = ToTime;
//										}
//									}
//								}
//								break;
//						}

//						hTime5 = PH_PY677_Time_Calc(STime, ETime);

//						////평일
//						if (DayOff == "1") {
//							////2교대 2조(야간조)
//							if (GNMUJO == "22") {
//								Base = Base - hTime5;
//								////기본근무
//								Midnight = Midnight - hTime5;
//								////심야시간에서 차감
//							////그외 주간조
//							} else {
//								Extend = Extend - hTime5;
//								////연장근무에서 차감
//								Midnight = Midnight - hTime5;
//								////심야시간에서 차감
//							}
//						} else {
//							////휴일
//							if (GNMUJO == "22") {
//								Special = Special - hTime5;
//								Midnight = Midnight - hTime5;
//								////심야시간에서 차감
//							} else {
//								////다음날 퇴근은 연장근무임
//								SpExtend = SpExtend - hTime5;
//								////연장근무에서 차감
//								Midnight = Midnight - hTime5;
//								////심야시간에서 차감
//							}
//						}
//						break;

//				}
//				oRecordSet.MoveNext();
//			}

//			oDS_PH_PY677B.SetValue("U_ColQty06", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(EarlyTo)));
//			//조출
//			oDS_PH_PY677B.SetValue("U_ColQty01", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Base, WorkType)));
//			//기본
//			oDS_PH_PY677B.SetValue("U_ColQty02", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Extend)));
//			//연장
//			oDS_PH_PY677B.SetValue("U_ColQty03", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Midnight)));
//			//심야

//			oDS_PH_PY677B.SetValue("U_ColQty07", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(SEarlyTo)));
//			//휴조
//			oDS_PH_PY677B.SetValue("U_ColQty04", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(Special)));
//			//특근
//			oDS_PH_PY677B.SetValue("U_ColQty05", pRow - 1, Convert.ToString(PH_PY677_hhmm_Calc(SpExtend)));
//			//특근연장

//			//    oForm.Items("EarlyTo").Specific.VALUE = PH_PY677_hhmm_Calc(EarlyTo) '//조출
//			//    oForm.Items("Base").Specific.VALUE = PH_PY677_hhmm_Calc(Base) '//기본
//			//    oForm.Items("Extend").Specific.VALUE = PH_PY677_hhmm_Calc(Extend) '//연장
//			//    oForm.Items("Midnight").Specific.VALUE = PH_PY677_hhmm_Calc(Midnight) '//심야
//			//
//			//    oForm.Items("SEarlyTo").Specific.VALUE = PH_PY677_hhmm_Calc(SEarlyTo) '//특근조출
//			//    oForm.Items("Special").Specific.VALUE = PH_PY677_hhmm_Calc(Special) '//특근
//			//    oForm.Items("SpExtend").Specific.VALUE = PH_PY677_hhmm_Calc(SpExtend) '//특근연장

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY677_Time_Calc_Main_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY677_Time_Calc_Main_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_Time_Calc_Main_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private double PH_PY677_hhmm_Calc(double mTime, string pWorkType = "")
//		{
//			double functionReturnValue = 0;

//			int hh = 0;
//			double MM = 0;

//			double Ret = 0;


//			hh = Conversion.Int((mTime) / 3600);
//			//UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
//			MM = ((mTime) % 3600) / 60;

//			if (MM > 0) {
//				if (MM > 30) {
//					MM = 1;
//				} else {
//					MM = 0.5;
//				}
//			} else if (MM == 0) {
//				MM = 0;
//			} else {
//				if (MM < -30) {
//					MM = -1;
//				} else {
//					MM = -0.5;
//				}
//			}

//			Ret = hh + MM;
//			//근태구분이 반차일 경우 무조건 4시간 반환(2014.04.21 송명규 추가) 별로 안 좋은 방법인데...
//			if (pWorkType == "D09" | pWorkType == "D10") {
//				Ret = 4;
//			}

//			//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
//			if (Information.IsDBNull(Ret)) {
//				Ret = 0;
//			}

//			functionReturnValue = Ret;
//			return functionReturnValue;
//			PH_PY677_hhmm_Calc_Exit:
//			return functionReturnValue;
//			PH_PY677_hhmm_Calc_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_hhmm_Calc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private double PH_PY677_Time_Calc(string GetTime, string OffTime)
//		{
//			double functionReturnValue = 0;

//			int STime = 0;
//			int ETime = 0;

//			int hh = 0;
//			double MM = 0;

//			double Ret = 0;


//			STime = Conversion.Int(Convert.ToDouble(Strings.Mid(GetTime, 1, 2))) * 3600 + Conversion.Int(Convert.ToDouble(Strings.Mid(GetTime, 3, 2))) * 60;

//			ETime = Conversion.Int(Convert.ToDouble(Strings.Mid(OffTime, 1, 2))) * 3600 + Conversion.Int(Convert.ToDouble(Strings.Mid(OffTime, 3, 2))) * 60;

//			//    hh = Int((ETime - STime) / 3600)
//			//    mm = ((ETime - STime) Mod 3600) / 60
//			//
//			//    If mm > 0 Then
//			//        If mm > 30 Then
//			//            mm = 1
//			//        Else
//			//            mm = 0.5
//			//        End If
//			//    ElseIf mm = 0 Then
//			//        mm = 0
//			//    Else
//			//        If mm < -30 Then
//			//            mm = -1
//			//        Else
//			//            mm = -0.5
//			//        End If
//			//    End If
//			//
//			//    Ret = hh + mm

//			Ret = ETime - STime;

//			//UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
//			if (Information.IsDBNull(Ret)) {
//				Ret = 0;
//			}

//			functionReturnValue = Ret;
//			return functionReturnValue;
//			PH_PY677_Time_Calct_Exit:
//			return functionReturnValue;
//			PH_PY677_Mark_Set_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_Time_Calc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY677_CheckAll()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string CheckType = null;
//			short loopCount = 0;

//			oForm.Freeze(true);
//			CheckType = "Y";

//			oMat01.FlushToDataSource();

//			for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++) {
//				if (Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg20", loopCount)) == "N") {
//					CheckType = "N";
//					break; // TODO: might not be correct. Was : Exit For
//				}
//			}

//			for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++) {
//				oDS_PH_PY677B.Offset = loopCount;
//				if (CheckType == "N") {
//					oDS_PH_PY677B.SetValue("U_ColReg20", loopCount, "Y");
//				} else {
//					oDS_PH_PY677B.SetValue("U_ColReg20", loopCount, "N");
//				}
//			}

//			oMat01.LoadFromDataSource();

//			oForm.Freeze(false);

//			return;
//			PH_PY677_CheckAll_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_CheckAll_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY677_ConfirmAll()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string ConfirmType = null;
//			short loopCount = 0;

//			oForm.Freeze(true);
//			ConfirmType = "Y";

//			oMat01.FlushToDataSource();

//			for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++) {
//				if (Strings.Trim(oDS_PH_PY677B.GetValue("U_ColReg16", loopCount)) == "N") {
//					ConfirmType = "N";
//					break; // TODO: might not be correct. Was : Exit For
//				}
//			}

//			for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++) {
//				oDS_PH_PY677B.Offset = loopCount;
//				if (ConfirmType == "N") {
//					oDS_PH_PY677B.SetValue("U_ColReg16", loopCount, "Y");
//				} else {
//					oDS_PH_PY677B.SetValue("U_ColReg16", loopCount, "N");
//				}
//			}

//			oMat01.LoadFromDataSource();

//			oForm.Freeze(false);

//			return;
//			PH_PY677_ConfirmAll_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY677_ConfirmAll_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

////Private Sub PH_PY677_Print_Report01()
////
////    Dim DocNum As String
////    Dim ErrNum As Integer
////    Dim WinTitle As String
////    Dim ReportName As String
////    Dim sQry As String
////    Dim oRecordSet As SAPbobsCOM.Recordset
////
////    On Error GoTo PH_PY677_Print_Report01_Error
////
////    Dim CLTCOD As String
////    Dim DestNo1 As String
////    Dim DestNo2 As String
////
////    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
////
////     '/ ODBC 연결 체크
////    If ConnectODBC = False Then
////        GoTo PH_PY677_Print_Report01_Error
////    End If
////
////    '//인자 MOVE , Trim 시키기..
////    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
////    DestNo1 = Trim(oForm.Items("DestNo1").Specific.VALUE)
////    DestNo2 = Trim(oForm.Items("DestNo2").Specific.VALUE)
////
////    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
////
////    WinTitle = "[PH_PY677] 공용증"
////
////    If CLTCOD = "1" Then '창원
////        ReportName = "PH_PY677_01.rpt"
////    ElseIf CLTCOD = "2" Then '동래
////        ReportName = "PH_PY677_02.rpt"
////    ElseIf CLTCOD = "3" Then '사상
////        ReportName = "PH_PY677_03.rpt"
////    End If
////
////    '// Formula 수식필드
////    ReDim gRpt_Formula(2)
////    ReDim gRpt_Formula_Value(2)
////
////    '// SubReport
////    ReDim gRpt_SRptSqry(1)
////    ReDim gRpt_SRptName(1)
////
////    ReDim gRpt_SFormula(1, 1)
////    ReDim gRpt_SFormula_Value(1, 1)
////
////    gRpt_SFormula(1, 1) = ""
////    gRpt_SFormula_Value(1, 1) = ""
////
////    '/ Procedure 실행"
////    sQry = "EXEC [PH_PY677_90] '" & CLTCOD & "','" & DestNo1 & "','" & DestNo2 & "'"
////
////    oRecordSet.DoQuery sQry
////    If oRecordSet.RecordCount = 0 Then
////        ErrNum = 1
////        GoTo PH_PY677_Print_Report01_Error
////    End If
////
////    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
////        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
////    End If
////
////    Set oRecordSet = Nothing
////    Exit Sub
////
////PH_PY677_Print_Report01_Error:
////    If ErrNum = 1 Then
////        Set oRecordSet = Nothing
////        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
////    Else
////    Set oRecordSet = Nothing
////    Sbo_Application.SetStatusBarMessage "PH_PY677_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
////    End If
////
////End Sub
//	}
//}
