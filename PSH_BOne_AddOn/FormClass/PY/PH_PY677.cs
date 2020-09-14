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
    }
}
