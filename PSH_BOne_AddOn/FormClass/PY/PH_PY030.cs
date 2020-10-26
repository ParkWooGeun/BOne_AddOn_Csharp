using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 공용등록
    /// </summary>
    internal class PH_PY030 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY030A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY030B;

        private string oLastItemUID01; // 클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01;  // 마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01;     // 마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Mode;        // 마지막 모드

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY030.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY030_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY030");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                //oForm.DataBrowser.BrowseBy = "Code";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY030_CreateItems();
                PH_PY030_ComboBox_Setting();
                PH_PY030_EnableMenus();
                PH_PY030_SetDocument(oFormDocEntry01);
                PH_PY030_FormResize();

                PH_PY030_LoadCaption();
                PH_PY030_FormItemEnabled();

                oForm.EnableMenu("1283", false);                // 삭제
                oForm.EnableMenu("1286", false);                // 닫기
                oForm.EnableMenu("1287", false);                // 복제
                oForm.EnableMenu("1285", false);                // 복원
                oForm.EnableMenu("1284", false);                // 취소
                oForm.EnableMenu("1293", false);                // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);

                string sQry = string.Empty;
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY030A]";
                oRecordSet.DoQuery(sQry);
                if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PH_PY030A.SetValue("DocEntry", 0, Convert.ToString(1));
                }
                else
                {
                    oDS_PH_PY030A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
                }
                oMat01.Columns.Item("Check").Visible = false; // 선택 체크박스 Visible = False
                PH_PY030_FormReset();
                //폼초기화 추가(2013.01.29 송명규)
                oForm.Update();

                // 기간
                oForm.Items.Item("SFrDate").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
                oForm.Items.Item("SToDate").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
                // 사번 포커스
                oForm.Items.Item("MSTCOD").Click();

            }
            catch (Exception ex)
            {
                oForm.Update();
                oForm.Freeze(false);
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
        /// <returns></returns>
        private void PH_PY030_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY030A = oForm.DataSources.DBDataSources.Item("@PH_PY030A");
                oDS_PH_PY030B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 관리번호
                oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

                // 사업장
                oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

                // 출장번호1
                oForm.DataSources.UserDataSources.Add("SDestNo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestNo1").Specific.DataBind.SetBound(true, "", "SDestNo1");

                // 출장번호2
                oForm.DataSources.UserDataSources.Add("SDestNo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestNo2").Specific.DataBind.SetBound(true, "", "SDestNo2");

                // 사원번호
                oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

                // 사원성명
                oForm.DataSources.UserDataSources.Add("SMSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SMSTNAM").Specific.DataBind.SetBound(true, "", "SMSTNAM");

                // 출장지
                oForm.DataSources.UserDataSources.Add("SDestinat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestinat").Specific.DataBind.SetBound(true, "", "SDestinat");

                // 출장지상세
                oForm.DataSources.UserDataSources.Add("SDest2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SDest2").Specific.DataBind.SetBound(true, "", "SDest2");

                // 작번
                oForm.DataSources.UserDataSources.Add("SCoCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SCoCode").Specific.DataBind.SetBound(true, "", "SCoCode");

                // 시작월
                oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

                // 종료월
                oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

                // 목적
                oForm.DataSources.UserDataSources.Add("SObject", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
                oForm.Items.Item("SObject").Specific.DataBind.SetBound(true, "", "SObject");

                // 비고
                oForm.DataSources.UserDataSources.Add("SComments", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100);
                oForm.Items.Item("SComments").Specific.DataBind.SetBound(true, "", "SComments");

                // 등록구분
                oForm.DataSources.UserDataSources.Add("SRegCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SRegCls").Specific.DataBind.SetBound(true, "", "SRegCls");

                // 목적구분
                oForm.DataSources.UserDataSources.Add("SObjCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SObjCls").Specific.DataBind.SetBound(true, "", "SObjCls");

                // 출장지역
                oForm.DataSources.UserDataSources.Add("SDestCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestCode").Specific.DataBind.SetBound(true, "", "SDestCode");

                // 출장구분
                oForm.DataSources.UserDataSources.Add("SDestDiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestDiv").Specific.DataBind.SetBound(true, "", "SDestDiv");

                // 차량구분
                oForm.DataSources.UserDataSources.Add("SVehicle", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SVehicle").Specific.DataBind.SetBound(true, "", "SVehicle");

                // 팀
                oForm.DataSources.UserDataSources.Add("STeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STeamCode").Specific.DataBind.SetBound(true, "", "STeamCode");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY030_LoadCaption()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_LoadCaption_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY030_Add_MatrixRow
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY030_Add_MatrixRow(int oRow, bool RowIserted = false)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (RowIserted == false)
                {
                    oDS_PH_PY030B.InsertRecord((oRow));
                }

                oMat01.AddRow();
                oDS_PH_PY030B.Offset = oRow;
                oDS_PH_PY030B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY030_MTX01
        /// </summary>
        private void PH_PY030_MTX01()
        {
            int i = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string sDocEntry = string.Empty;    // 관리번호
            string sCLTCOD = string.Empty;      // 사업장
            string SDestNo1 = string.Empty;     // 출장번호1
            string SDestNo2 = string.Empty;     // 출장번호2
            string sMSTCOD = string.Empty;      // 사원번호
            string SDestinat = string.Empty;    // 출장지
            string SDest2 = string.Empty;       // 출장지상세
            string SCoCode = string.Empty;      // 작번
            string SFrDate = string.Empty;      // 시작일자
            string SToDate = string.Empty;      // 종료일자
            string SObject = string.Empty;      // 목적
            string SComments = string.Empty;    // 비고
            string SRegCls = string.Empty;      // 등록구분
            string SObjCls = string.Empty;      // 목적구분
            string SDestCode = string.Empty;    // 출장지역
            string SDestDiv = string.Empty;     // 출장구분
            string SVehicle = string.Empty;     // 차량구분
            string sTeamCode = string.Empty;    // 팀(2013.06.01 송명규 추가)

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);
            try
            {
                oForm.Freeze(true);

                sDocEntry = oForm.Items.Item("SDocEntry").Specific.VALUE.ToString().Trim();                // 관리번호
                sCLTCOD   = oForm.Items.Item("SCLTCOD").Specific.VALUE.ToString().Trim();                  // 사업장
                SDestNo1  = oForm.Items.Item("SDestNo1").Specific.VALUE.ToString().Trim();                 // 출장번호1
                SDestNo2  = oForm.Items.Item("SDestNo2").Specific.VALUE.ToString().Trim();                 // 출장번호2
                sMSTCOD   = oForm.Items.Item("SMSTCOD").Specific.VALUE.ToString().Trim();                  // 사원번호
                SDestinat = oForm.Items.Item("SDestinat").Specific.VALUE.ToString().Trim();                // 출장지
                SDest2    = oForm.Items.Item("SDest2").Specific.VALUE.ToString().Trim();                   // 출장지상세
                SCoCode   = oForm.Items.Item("SCoCode").Specific.VALUE.ToString().Trim();                  // 작번
                SFrDate   = oForm.Items.Item("SFrDate").Specific.VALUE.ToString().Trim().Replace(".", ""); // 시작일자
                SToDate   = oForm.Items.Item("SToDate").Specific.VALUE.ToString().Trim().Replace(".", ""); // 종료일자
                SObject   = oForm.Items.Item("SObject").Specific.VALUE.ToString().Trim();                  // 목적
                SComments = oForm.Items.Item("SComments").Specific.VALUE.ToString().Trim();                // 비고
                SRegCls   = oForm.Items.Item("SRegCls").Specific.VALUE.ToString().Trim();                  // 등록구분
                SObjCls   = oForm.Items.Item("SObjCls").Specific.VALUE.ToString().Trim();                  // 목적구분
                SDestCode = oForm.Items.Item("SDestCode").Specific.VALUE.ToString().Trim();                // 출장지역
                SDestDiv  = oForm.Items.Item("SDestDiv").Specific.VALUE.ToString().Trim();                 // 출장구분
                SVehicle  = oForm.Items.Item("SVehicle").Specific.VALUE.ToString().Trim();                 // 차량구분
                sTeamCode = oForm.Items.Item("STeamCode").Specific.VALUE.ToString().Trim();                // 팀

                sQry = "            EXEC [PH_PY030_01] ";
                sQry = sQry + "'" + sDocEntry + "',";
                sQry = sQry + "'" + sCLTCOD + "',";
                sQry = sQry + "'" + SDestNo1 + "',";
                sQry = sQry + "'" + SDestNo2 + "',";
                sQry = sQry + "'" + sMSTCOD + "',";
                sQry = sQry + "'" + SDestinat + "',";
                sQry = sQry + "'" + SDest2 + "',";
                sQry = sQry + "'" + SCoCode + "',";
                sQry = sQry + "'" + SFrDate + "',";
                sQry = sQry + "'" + SToDate + "',";
                sQry = sQry + "'" + SObject + "',";
                sQry = sQry + "'" + SComments + "',";
                sQry = sQry + "'" + SRegCls + "',";
                sQry = sQry + "'" + SObjCls + "',";
                sQry = sQry + "'" + SDestCode + "',";
                sQry = sQry + "'" + SDestDiv + "',";
                sQry = sQry + "'" + SVehicle + "',";
                sQry = sQry + "'" + sTeamCode + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY030B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if ((oRecordSet01.RecordCount == 0))
                {
                    ErrNum = 1;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY030_LoadCaption();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY030B.Size)
                    {
                        oDS_PH_PY030B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY030B.Offset = i;

                    oDS_PH_PY030B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY030B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());   // 관리번호
                    oDS_PH_PY030B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim());     // 사업장
                    oDS_PH_PY030B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("DestNo1").Value.ToString().Trim());    // 출장번호1
                    oDS_PH_PY030B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim());    // 출장번호2
                    oDS_PH_PY030B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());     // 사원번호
                    oDS_PH_PY030B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim());     // 사원성명
                    oDS_PH_PY030B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("Destinat").Value.ToString().Trim());   // 출장지
                    oDS_PH_PY030B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("Dest2").Value.ToString().Trim());      // 출장지상세
                    oDS_PH_PY030B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("CoCode").Value.ToString().Trim());     // 작번
                    oDS_PH_PY030B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("FrDate").Value.ToString().Trim());     // 시작일자
                    oDS_PH_PY030B.SetValue("U_ColTm01", i,  oRecordSet01.Fields.Item("FrTime").Value.ToString().Trim());     // 시작시각
                    oDS_PH_PY030B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("ToDate").Value.ToString().Trim());     // 종료일자
                    oDS_PH_PY030B.SetValue("U_ColTm02", i,  oRecordSet01.Fields.Item("ToTime").Value.ToString().Trim());     // 종료시각
                    oDS_PH_PY030B.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("Object").Value.ToString().Trim());     // 목적
                    oDS_PH_PY030B.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim());   // 비고
                    oDS_PH_PY030B.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("RegCls").Value.ToString().Trim());     // 등록구분
                    oDS_PH_PY030B.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("ObjCls").Value.ToString().Trim());     // 목적구분
                    oDS_PH_PY030B.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("DestCode").Value.ToString().Trim());   // 출장지역
                    oDS_PH_PY030B.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("DestDiv").Value.ToString().Trim());    // 출장구분
                    oDS_PH_PY030B.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("Vehicle").Value.ToString().Trim());    // 차량구분
                    oDS_PH_PY030B.SetValue("U_ColPrc01", i, oRecordSet01.Fields.Item("FuelPrc").Value.ToString().Trim());    // 1L단가
                    oDS_PH_PY030B.SetValue("U_ColReg22", i, oRecordSet01.Fields.Item("FuelType").Value.ToString().Trim());   // 유류
                    oDS_PH_PY030B.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("Distance").Value.ToString().Trim());   // 거리
                    oDS_PH_PY030B.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("TransExp").Value.ToString().Trim());   // 교통비
                    oDS_PH_PY030B.SetValue("U_ColSum02", i, oRecordSet01.Fields.Item("DayExp").Value.ToString().Trim());     // 일비
                    oDS_PH_PY030B.SetValue("U_ColReg23", i, oRecordSet01.Fields.Item("FoodNum").Value.ToString().Trim());    // 식수
                    oDS_PH_PY030B.SetValue("U_ColSum03", i, oRecordSet01.Fields.Item("FoodExp").Value.ToString().Trim());    // 식비
                    oDS_PH_PY030B.SetValue("U_ColSum04", i, oRecordSet01.Fields.Item("ParkExp").Value.ToString().Trim());    // 주차비
                    oDS_PH_PY030B.SetValue("U_ColSum05", i, oRecordSet01.Fields.Item("TollExp").Value.ToString().Trim());    // 도로비
                    oDS_PH_PY030B.SetValue("U_ColSum06", i, oRecordSet01.Fields.Item("TotalExp").Value.ToString().Trim());   // 합계

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY030_DeleteData
        /// </summary>
        private void PH_PY030_DeleteData()
        {
            int ErrNum = 0;
            string sQry = string.Empty; ;
            string DocEntry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {

                    DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PH_PY030A] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if ((oRecordSet01.RecordCount == 0))
                    {
                        ErrNum = 1;
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "EXEC PH_PY030_04 '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }

                dataHelpClass.MDC_GF_Message("삭제 완료!", "S");
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("삭제대상이 없습니다. 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_DeleteData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY030_UpdateData
        /// </summary>
        /// <returns></returns>
        private bool PH_PY030_UpdateData()
        {
            bool functionReturnValue = false;

            string sQry = String.Empty;
            short DocEntry = 0;                      // 관리번호
            string CLTCOD = String.Empty;            // 사업장
            string DestNo1 = String.Empty;           // 출장번호1
            string DestNo2 = String.Empty;           // 출장번호2
            string MSTCOD = String.Empty;            // 사원번호
            string MSTNAM = String.Empty;            // 사원성명
            string Destinat = String.Empty;          // 출장지
            string Dest2 = String.Empty;             // 출장지상세
            string CoCode = String.Empty;            // 작번
            string FrDate = String.Empty;            // 시작일자
            string FrTime = String.Empty;            // 시작시각
            string ToDate = String.Empty;            // 종료일자
            string ToTime = String.Empty;            // 종료시각
            string Object_Renamed = String.Empty;    // 목적
            string Comments = String.Empty;          // 비고
            string RegCls = String.Empty;            // 등록구분
            string ObjCls = String.Empty;            // 목적구분
            string DestCode = String.Empty;          // 출장지역
            string DestDiv = String.Empty;           // 출장구분
            string Vehicle = String.Empty;           // 차량구분
            double FuelPrc = 0;                      // 1L단가
            string FuelType = String.Empty;          // 유류
            double Distance = 0;                     // 거리
            double TransExp = 0;                     // 교통비
            double DayExp = 0;                       // 일비
            string FoodNum = String.Empty;           // 식수
            double FoodExp = 0;                      // 식비
            double ParkExp = 0;                      // 주차비
            double TollExp = 0;                      // 도로비
            double TotalExp = 0;                     // 합계

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = Convert.ToInt16(oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim());
                CLTCOD   = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                DestNo1  = oForm.Items.Item("DestNo1").Specific.VALUE.ToString().Trim();
                DestNo2  = oForm.Items.Item("DestNo2").Specific.VALUE.ToString().Trim();
                MSTCOD   = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                MSTNAM   = oForm.Items.Item("MSTNAM").Specific.VALUE.ToString().Trim();
                Destinat = oForm.Items.Item("Destinat").Specific.VALUE.ToString().Trim();
                Dest2    = oForm.Items.Item("Dest2").Specific.VALUE.ToString().Trim();
                CoCode   = oForm.Items.Item("CoCode").Specific.VALUE.ToString().Trim();
                FrDate   = oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim();
                FrTime   = oForm.Items.Item("FrTime").Specific.VALUE.ToString().Trim();
                ToDate   = oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim();
                ToTime   = oForm.Items.Item("ToTime").Specific.VALUE.ToString().Trim();
                Object_Renamed = oForm.Items.Item("Object").Specific.VALUE.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.VALUE.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.VALUE.ToString().Trim();
                ObjCls   = oForm.Items.Item("ObjCls").Specific.VALUE.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.VALUE.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.VALUE.ToString().Trim();
                Vehicle  = oForm.Items.Item("Vehicle").Specific.VALUE.ToString().Trim();
                FuelPrc  = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.VALUE.ToString().Trim());
                FuelType = oForm.Items.Item("FuelType").Specific.VALUE.ToString().Trim();
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.VALUE.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.VALUE.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.VALUE.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.VALUE.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.VALUE.ToString().Trim());
                ParkExp  = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.VALUE.ToString().Trim());
                TollExp  = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.VALUE.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.VALUE.ToString().Trim());


                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).ToString().Trim()))
                {
                    dataHelpClass.MDC_GF_Message("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E");
                    functionReturnValue = false;
                    throw new Exception();
                }

                sQry = "            EXEC [PH_PY030_03] ";
                sQry = sQry + "'" + DocEntry + "',";              // 관리번호
                sQry = sQry + "'" + CLTCOD + "',";                // 사업장
                sQry = sQry + "'" + DestNo1 + "',";               // 출장번호1
                sQry = sQry + "'" + DestNo2 + "',";               // 출장번호2
                sQry = sQry + "'" + MSTCOD + "',";                // 사원번호
                sQry = sQry + "'" + MSTNAM + "',";                // 사원성명
                sQry = sQry + "'" + Destinat + "',";              // 출장지
                sQry = sQry + "'" + Dest2 + "',";                 // 출장지상세
                sQry = sQry + "'" + CoCode + "',";                // 작번
                sQry = sQry + "'" + FrDate + "',";                // 시작일자
                sQry = sQry + "'" + FrTime + "',";                // 시작시각
                sQry = sQry + "'" + ToDate + "',";                // 종료일자
                sQry = sQry + "'" + ToTime + "',";                // 종료시각
                sQry = sQry + "'" + Object_Renamed + "',";        // 목적
                sQry = sQry + "'" + Comments + "',";              // 비고
                sQry = sQry + "'" + RegCls + "',";                // 등록구분
                sQry = sQry + "'" + ObjCls + "',";                // 목적구분
                sQry = sQry + "'" + DestCode + "',";              // 출장지역
                sQry = sQry + "'" + DestDiv + "',";               // 출장구분
                sQry = sQry + "'" + Vehicle + "',";               // 차량구분
                sQry = sQry + "'" + FuelPrc + "',";               // 1L단가
                sQry = sQry + "'" + FuelType + "',";              // 유류
                sQry = sQry + "'" + Distance + "',";              // 거리
                sQry = sQry + "'" + TransExp + "',";              // 교통비
                sQry = sQry + "'" + DayExp + "',";                // 일비
                sQry = sQry + "'" + FoodNum + "',";               // 식수
                sQry = sQry + "'" + FoodExp + "',";               // 식비
                sQry = sQry + "'" + ParkExp + "',";               // 주차비
                sQry = sQry + "'" + TollExp + "',";               // 도로비
                sQry = sQry + "'" + TotalExp + "'";               // 합계

                oRecordSet01.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("수정 완료!", "S");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_UpdateData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                functionReturnValue = true;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY030_DeleteData
        /// </summary>
        private bool PH_PY030_AddData()
        {
            bool functionReturnValue = false;
            int DocEntry = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;          // 사업장
            string DestNo1 = string.Empty;         // 출장번호1
            string DestNo2 = string.Empty;         // 출장번호2
            string MSTCOD = string.Empty;          // 사원번호
            string MSTNAM = string.Empty;          // 사원성명
            string Destinat = string.Empty;        // 출장지
            string Dest2 = string.Empty;           // 출장지상세
            string CoCode = string.Empty;          // 작번
            string FrDate = string.Empty;          // 시작일자
            string FrTime = string.Empty;          // 시작시각
            string ToDate = string.Empty;          // 종료일자
            string ToTime = string.Empty;          // 종료시각
            string Object_Renamed = string.Empty;  // 목적
            string Comments = string.Empty;        // 비고
            string RegCls = string.Empty;          // 등록구분
            string ObjCls = string.Empty;          // 목적구분
            string DestCode = string.Empty;        // 출장지역
            string DestDiv = string.Empty;         // 출장구분
            string Vehicle = string.Empty;         // 차량구분
            double FuelPrc = 0;                    // 1L단가
            string FuelType = string.Empty;        // 유류
            double Distance = 0;                   // 거리
            double TransExp = 0;                   // 교통비
            double DayExp = 0;                     // 일비
            string FoodNum = string.Empty;         // 식수
            double FoodExp = 0;                    // 식비
            double ParkExp = 0;                    // 주차비
            double TollExp = 0;                    // 도로비
            double TotalExp = 0;                   // 합계
            string UserSign = string.Empty;        // UserSign

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                CLTCOD   = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                DestNo1  = oForm.Items.Item("DestNo1").Specific.VALUE.ToString().Trim();
                DestNo2  = oForm.Items.Item("DestNo2").Specific.VALUE.ToString().Trim();
                MSTCOD   = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                MSTNAM   = oForm.Items.Item("MSTNAM").Specific.VALUE.ToString().Trim();
                Destinat = oForm.Items.Item("Destinat").Specific.VALUE.ToString().Trim();
                Dest2    = oForm.Items.Item("Dest2").Specific.VALUE.ToString().Trim();
                CoCode   = oForm.Items.Item("CoCode").Specific.VALUE.ToString().Trim();
                FrDate   = oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim();
                FrTime   = oForm.Items.Item("FrTime").Specific.VALUE.ToString().Trim();
                ToDate   = oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim();
                ToTime   = oForm.Items.Item("ToTime").Specific.VALUE.ToString().Trim();
                Object_Renamed = oForm.Items.Item("Object").Specific.VALUE.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.VALUE.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.VALUE.ToString().Trim();
                ObjCls   = oForm.Items.Item("ObjCls").Specific.VALUE.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.VALUE.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.VALUE.ToString().Trim();
                Vehicle  = oForm.Items.Item("Vehicle").Specific.VALUE.ToString().Trim();
                FuelPrc  = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.VALUE.ToString().Trim());
                FuelType = oForm.Items.Item("FuelType").Specific.VALUE.ToString().Trim();
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.VALUE.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.VALUE.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.VALUE.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.VALUE.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.VALUE.ToString().Trim());
                ParkExp  = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.VALUE.ToString().Trim());
                TollExp  = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.VALUE.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.VALUE.ToString().Trim());
                UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY030A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }

                sQry = "            EXEC [PH_PY030_02] ";
                sQry = sQry + "'" + DocEntry + "',";       // 관리번호
                sQry = sQry + "'" + CLTCOD + "',";         // 사업장
                sQry = sQry + "'" + DestNo1 + "',";        // 출장번호1
                sQry = sQry + "'" + DestNo2 + "',";        // 출장번호2
                sQry = sQry + "'" + MSTCOD + "',";         // 사원번호
                sQry = sQry + "'" + MSTNAM + "',";         // 사원성명
                sQry = sQry + "'" + Destinat + "',";       // 출장지
                sQry = sQry + "'" + Dest2 + "',";          // 출장지상세
                sQry = sQry + "'" + CoCode + "',";         // 작번
                sQry = sQry + "'" + FrDate + "',";         // 시작일자
                sQry = sQry + "'" + FrTime + "',";         // 시작시각
                sQry = sQry + "'" + ToDate + "',";         // 종료일자
                sQry = sQry + "'" + ToTime + "',";         // 종료시각
                sQry = sQry + "'" + Object_Renamed + "',"; // 목적
                sQry = sQry + "'" + Comments + "',";       // 비고
                sQry = sQry + "'" + RegCls + "',";         // 등록구분
                sQry = sQry + "'" + ObjCls + "',";         // 목적구분
                sQry = sQry + "'" + DestCode + "',";       // 출장지역
                sQry = sQry + "'" + DestDiv + "',";        // 출장구분
                sQry = sQry + "'" + Vehicle + "',";        // 차량구분
                sQry = sQry + "'" + FuelPrc + "',";        // 1L단가
                sQry = sQry + "'" + FuelType + "',";       // 유류
                sQry = sQry + "'" + Distance + "',";       // 거리
                sQry = sQry + "'" + TransExp + "',";       // 교통비
                sQry = sQry + "'" + DayExp + "',";         // 일비
                sQry = sQry + "'" + FoodNum + "',";        // 식수
                sQry = sQry + "'" + FoodExp + "',";        // 식비
                sQry = sQry + "'" + ParkExp + "',";        // 주차비
                sQry = sQry + "'" + TollExp + "',";        // 도로비
                sQry = sQry + "'" + TotalExp + "',";       // 합계
                sQry = sQry + "'" + UserSign + "'";        // UserSign

                oRecordSet02.DoQuery(sQry);

                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_AddData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                functionReturnValue = true;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PH_PY030_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            int ErrNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("DestNo1").Specific.VALUE.ToString().Trim()))      // 출장번호1
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DestNo2").Specific.VALUE.ToString().Trim()))  // 출장번호2
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim()))  // 사원번호
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim()))  // 시작일자
                {
                    ErrNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FrTime").Specific.VALUE.ToString().Trim()))  // 시작시각
                {
                    ErrNum = 5;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.VALUE.ToString().Trim()))  // 종료일자
                {
                    ErrNum = 6;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToTime").Specific.VALUE.ToString().Trim()))  // 종료시각
                {
                    ErrNum = 7;
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("출장번호1은 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("DestNo1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 2)
                {
                    dataHelpClass.MDC_GF_Message("출장번호2는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("DestNo2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 3)
                {
                    dataHelpClass.MDC_GF_Message("사원번호는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 4)
                {
                    dataHelpClass.MDC_GF_Message("시작일자는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 5)
                {
                    dataHelpClass.MDC_GF_Message("시작시각은 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 6)
                {
                    dataHelpClass.MDC_GF_Message("종료일자는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 7)
                {
                    dataHelpClass.MDC_GF_Message("종료시각은 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
                functionReturnValue = true;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY030_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
        {
            int loopCount = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sCLTCOD = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                switch (oUID)
                {
                    case "FuelPrc":
                        PH_PY030_CalculateTransExp();
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "Distance":
                        PH_PY030_CalculateTransExp();
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "TransExp":
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "DayExp":
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "FoodExp":
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "ParkExp":
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "TollExp":
                        PH_PY030_CalculateTotalExp();
                        break;
                    case "FrDate":
                        PH_PY030_GetDestNo();
                        break;
                    case "FuelType":
                        PH_PY030_GetFuelPrc();
                        break;
                    case "CLTCOD":
                        PH_PY030_GetDestNo();                        // 출장번호생성
                        break;
                    case "FoodNum":
                        PH_PY030_CalculateFoodExp();                 // 식비 계산
                        PH_PY030_CalculateTotalExp();
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
                        sQry = "        SELECT      U_Code,";
                        sQry = sQry + "             U_CodeNm";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '1'";
                        sQry = sQry + "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.VALUE + "'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList((oForm.Items.Item("STeamCode").Specific), sQry,"", false, false);
                        oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        oForm.Items.Item("STeamCode").DisplayDesc = true;
                        break;
                    case "Destinat":
                        if (oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() == "1")
                        {
                            sQry = "Select Distance = U_Num2 From [@PS_HR200L] Where Code = 'P217' And U_Code = '" + oForm.Items.Item("Destinat").Specific.VALUE.ToString().Trim() + "' and U_Char1 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("Distance").Specific.VALUE = oRecordSet01.Fields.Item("Distance").Value;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY030_ComboBox_Setting
        /// </summary>
        public void PH_PY030_ComboBox_Setting()
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // 기본정보
                // 출장지
                oForm.Items.Item("Destinat").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P217'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Destinat").Specific, sQry,  "", false,  false);
                oForm.Items.Item("Destinat").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 등록구분
                oForm.Items.Item("RegCls").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P223'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("RegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("RegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 목적구분
                oForm.Items.Item("ObjCls").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P224'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("ObjCls").Specific, sQry, "", false, false);
                oForm.Items.Item("ObjCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("DestCode").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P225'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("DestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("DestDiv").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P216'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("DestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 차량구분
                oForm.Items.Item("Vehicle").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P218'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Vehicle").Specific, sQry, "", false, false);
                oForm.Items.Item("Vehicle").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 유류
                oForm.Items.Item("FuelType").Specific.ValidValues.Add("%", "선택");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P226'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("FuelType").Specific, sQry, "", false, false);
                oForm.Items.Item("FuelType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 식수
                //    Call oForm.Items("FoodNum").Specific.ValidValues.Add("0", "선택")
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P227'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("FoodNum").Specific, sQry, "", false, false);
                oForm.Items.Item("FoodNum").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 조회정보
                // 출장지
                oForm.Items.Item("SDestinat").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P217'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestinat").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestinat").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 등록구분
                oForm.Items.Item("SRegCls").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P223'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SRegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("SRegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 목적구분
                oForm.Items.Item("SObjCls").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P224'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SObjCls").Specific, sQry, "", false, false);
                oForm.Items.Item("SObjCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("SDestCode").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P225'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("SDestDiv").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P216'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 차량구분
                oForm.Items.Item("SVehicle").Specific.ValidValues.Add("%", "전체");
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P218'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SVehicle").Specific, sQry, "", false, false);
                oForm.Items.Item("SVehicle").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 매트릭스
                // 사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId","","");

                // 출장지
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P217'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Destinat"), sQry, "", "");

                // 등록구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P223'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RegCls"), sQry, "", "");

                // 목적구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P224'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ObjCls"), sQry, "", "");

                // 출장지역
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P225'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestCode"), sQry, "", "");

                // 출장구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P216'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestDiv"), sQry, "", "");

                // 차량구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P218'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Vehicle"), sQry, "", "");

                // 유류
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P226'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("FuelType"), sQry, "", "");

                // 식수
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm As [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P227'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("FoodNum"), sQry, "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_ComboBox_Setting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY030_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);               ////취소
                oForm.EnableMenu("1287", true);                ////복제
                oForm.EnableMenu("1293", false);               ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY030_SetDocument(string oFormDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry01)))
                {
                    PH_PY030_FormItemEnabled();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY030_FormResize
        /// </summary>
        private void PH_PY030_FormResize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_FormResize_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY030_FormReset
        /// </summary>
        public void PH_PY030_FormReset()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                //관리번호
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY030A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PH_PY030A.SetValue("DocEntry", 0, Convert.ToString(1));
                }
                else
                {
                    oDS_PH_PY030A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1));
                }

                string User_BPLID = null;
                User_BPLID = dataHelpClass.User_BPLID();

                // 기준정보
                oDS_PH_PY030A.SetValue("U_CLTCOD", 0, User_BPLID);                          // 사업장
                oDS_PH_PY030A.SetValue("U_DestNo1", 0, "");                                 // 출장번호1
                oDS_PH_PY030A.SetValue("U_DestNo2", 0, "");                                 // 출장번호2
                oDS_PH_PY030A.SetValue("U_MSTCOD", 0, "");                                  // 사원번호
                oDS_PH_PY030A.SetValue("U_MSTNAM", 0, "");                                  // 사원성명
                oDS_PH_PY030A.SetValue("U_Destinat", 0, "%");                               // 출장지
                oDS_PH_PY030A.SetValue("U_Dest2", 0, "");                                   // 출장지상세
                oDS_PH_PY030A.SetValue("U_CoCode", 0, "");                                  // 작번
                oDS_PH_PY030A.SetValue("U_FrDate", 0, DateTime.Now.ToString("yyyyMMdd"));   // 시작일자
                oDS_PH_PY030A.SetValue("U_FrTime", 0, "");                                  // 시작시각
                oDS_PH_PY030A.SetValue("U_ToDate", 0, DateTime.Now.ToString("yyyyMMdd"));   // 종료일자
                oDS_PH_PY030A.SetValue("U_ToTime", 0, "");                                  // 종료시각
                oDS_PH_PY030A.SetValue("U_Object", 0, "");                                  // 목적
                oDS_PH_PY030A.SetValue("U_Comments", 0, "");                                // 비고
                oDS_PH_PY030A.SetValue("U_RegCls", 0, "01");                                // 등록구분
                oDS_PH_PY030A.SetValue("U_ObjCls", 0, "%");                                 // 목적구분
                oDS_PH_PY030A.SetValue("U_DestCode", 0, "01");                              // 출장지역
                oDS_PH_PY030A.SetValue("U_DestDiv", 0, "%");                                // 출장구분
                oDS_PH_PY030A.SetValue("U_Vehicle", 0, "%");                                // 차량구분
                oDS_PH_PY030A.SetValue("U_FuelPrc", 0, Convert.ToString(0));                // 1L단가
                oDS_PH_PY030A.SetValue("U_FuelType", 0, "%");                               // 유류
                oDS_PH_PY030A.SetValue("U_Distance", 0, Convert.ToString(0));               // 거리
                oDS_PH_PY030A.SetValue("U_TransExp", 0, Convert.ToString(0));               // 교통비
                oDS_PH_PY030A.SetValue("U_DayExp", 0, Convert.ToString(0));                 // 일비
                oDS_PH_PY030A.SetValue("U_FoodNum", 0, "0");                                // 식수
                oDS_PH_PY030A.SetValue("U_FoodExp", 0, Convert.ToString(0));                // 식비
                oDS_PH_PY030A.SetValue("U_ParkExp", 0, Convert.ToString(0));                // 주차비
                oDS_PH_PY030A.SetValue("U_TollExp", 0, Convert.ToString(0));                // 도로비
                oDS_PH_PY030A.SetValue("U_TotalExp", 0, Convert.ToString(0));               // 합계
                // 출장번호
                PH_PY030_GetDestNo();

                oForm.Items.Item("MSTCOD").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY030_CalculateTransExp
        /// </summary>
        private void PH_PY030_CalculateTransExp()
        {
            try
            {
                double FuelPrc = 0;                //유류단가
                double Distance = 0;                //거리
                int TransExp = 0;                //교통비

                FuelPrc = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.VALUE.ToString().Trim());
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.VALUE.ToString().Trim());

                TransExp = Convert.ToInt32((FuelPrc * Distance * 0.1) / 10) * 10;                //원단위  절사 2020.06.03 황영수

                oForm.Items.Item("TransExp").Specific.VALUE = TransExp;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_CalculateTransExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY030_CalculateTransExp
        /// </summary>
        private void PH_PY030_CalculateTotalExp()
        {
            double TransExp = 0;               // 교통비
            double DayExp = 0;                 // 일비
            double FoodExp = 0;                // 식비
            double ParkExp = 0;                // 주차비
            double TollExp = 0;                // 도로비
            double TotalExp = 0;               // 합계
            try
            {
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.VALUE.ToString().Trim());
                DayExp = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.VALUE.ToString().Trim());
                FoodExp = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.VALUE.ToString().Trim());
                ParkExp = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.VALUE.ToString().Trim());
                TollExp = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.VALUE.ToString().Trim());
                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp;

                oDS_PH_PY030A.SetValue("U_TotalExp", 0, Convert.ToString(TotalExp));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_CalculateTotalExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY030_GetDestNo
        /// </summary>
        private void PH_PY030_GetDestNo()
        {
            string FrDate = string.Empty;
            string CLTCOD = string.Empty;
            string sQry = string.Empty;

            PSH_CodeHelpClass codeaHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                FrDate = codeaHelpClass.Left(oForm.Items.Item("FrDate").Specific.VALUE.ToString().Trim(), 6);

                sQry = "EXEC PH_PY030_05 '" + CLTCOD + "', '" + FrDate + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("DestNo1").Specific.VALUE = FrDate;
                oForm.Items.Item("DestNo2").Specific.VALUE = oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_GetDestNo_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY030_GetFuelPrc
        /// </summary>
        private void PH_PY030_GetFuelPrc()
        {
            string CLTCOD = string.Empty;
            string sQry = string.Empty;
            object CheckAmt = string.Empty;
            string StdYear = string.Empty;
            string StdMonth = string.Empty;
            string FuelType = string.Empty;
            double FuelPrice = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                CLTCOD = oDS_PH_PY030A.GetValue("U_CLTCOD", 0).ToString().Trim();                // 사업장

                if (!string.IsNullOrEmpty(oDS_PH_PY030A.GetValue("U_FrDate", 0).ToString().Trim()))
                {
                    StdYear  = oDS_PH_PY030A.GetValue("U_FrDate", 0).Substring(0, 4).ToString().Trim();
                    StdMonth = oDS_PH_PY030A.GetValue("U_FrDate", 0).Substring(4, 2).ToString().Trim();
                }

                FuelType = oDS_PH_PY030A.GetValue("U_FuelType", 0).ToString().Trim();            // 유류

                sQry = "           SELECT      T0.U_Year AS [StdYear],";
                sQry = sQry + "                T1.U_Month AS [StdMonth],";
                sQry = sQry + "                T1.U_Gasoline AS [Gasoline],";
                sQry = sQry + "                T1.U_Diesel AS [Diesel],";
                sQry = sQry + "                T1.U_LPG AS [LPG]";
                sQry = sQry + " FROM       [@PH_PY007A] AS T0";
                sQry = sQry + "                INNER JOIN";
                sQry = sQry + "                [@PH_PY007B] AS T1";
                sQry = sQry + "                    ON T0.Code = T1.Code";
                sQry = sQry + " WHERE      T0.U_CLTCOD = '" + CLTCOD + "'";
                sQry = sQry + "                AND T0.U_Year = '" + StdYear + "'";
                sQry = sQry + "                AND T1.U_Month = '" + StdMonth + "'";

                oRecordSet01.DoQuery(sQry);

                //휘발유
                if (FuelType == "1")
                {
                    FuelPrice = oRecordSet01.Fields.Item("Gasoline").Value;                    //가스
                }
                else if (FuelType == "2")
                {
                    FuelPrice = oRecordSet01.Fields.Item("LPG").Value;                    //경유
                }
                else if (FuelType == "3")
                {
                    FuelPrice = oRecordSet01.Fields.Item("Diesel").Value;
                }
                else
                {
                    FuelPrice = 0;
                }
                oDS_PH_PY030A.SetValue("U_FuelPrc", 0, Convert.ToString(FuelPrice));
                oForm.Items.Item("Distance").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_GetFuelPrc_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        private void PH_PY030_CalculateFoodExp()
        {
            short ErrNum = 0;
            string sQry = null;
            string MSTCOD = null;            // 사번
            short FoodNum = 0;               // 식수
            double FoodPrc = 0;              // 당일식비
            double FoodExp = 0;              // 전체식비
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // 사번을 선택하지 않으면
                if (string.IsNullOrEmpty(oDS_PH_PY030A.GetValue("U_MSTCOD", 0).ToString().Trim()) & oDS_PH_PY030A.GetValue("U_FoodNum", 0).ToString().Trim() != "0")
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                MSTCOD = oDS_PH_PY030A.GetValue("U_MSTCOD", 0).ToString().Trim();                       // 사번
                FoodNum = Convert.ToInt16(oDS_PH_PY030A.GetValue("U_FoodNum", 0).ToString().Trim());    // 식수

                sQry = "            SELECT      T1.U_Num4 AS [FoodPrc]";
                sQry = sQry + "  FROM       [@PH_PY001A] AS T0";
                sQry = sQry + "                 LEFT JOIN";
                sQry = sQry + "                 [@PS_HR200L] AS T1";
                sQry = sQry + "                     ON T0.U_JIGCOD = T1.U_Code";
                sQry = sQry + "                     AND T1.Code = 'P232'";
                sQry = sQry + "                     AND T1.U_UseYN = 'Y'";
                sQry = sQry + "  WHERE      T0.Code = '" + MSTCOD + "'";

                oRecordSet01.DoQuery(sQry);

                FoodPrc = oRecordSet01.Fields.Item("FoodPrc").Value;
                FoodExp = FoodPrc * FoodNum;

                oDS_PH_PY030A.SetValue("U_FoodExp", 0, Convert.ToString(FoodExp));
                oForm.Items.Item("FoodExp").Click();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사원을 먼저 선택하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_CalculateFoodExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY030_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", true);
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                    //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                    //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK:                    //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                    //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                    //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE:                    //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                    //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    //18
                //    break;
                //    et_FORM_ACTIVATE
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    //19
                //    break;
                //    et_FORM_DEACTIVATE
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                    ////20
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                    //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                    //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                    //4
                //    break;
                //    et_LOST_FOCUS
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "PH_PY030")
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

                    // 추가/확인 버튼클릭
                    if (pVal.ItemUID == "BtnAdd")
                    {

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {

                            if (PH_PY030_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY030_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY030_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY030_LoadCaption();
                            PH_PY030_MTX01();

                            oLast_Mode = (int)oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY030_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY030_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY030_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY030_LoadCaption();
                            PH_PY030_MTX01();
                        }

                        // 조회
                    }
                    else if (pVal.ItemUID == "BtnSearch")
                    {
                        PH_PY030_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        // fm_VIEW_MODE
                        PH_PY030_LoadCaption();
                        PH_PY030_MTX01();
                        // 삭제
                    }
                    else if (pVal.ItemUID == "BtnDelete")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            PH_PY030_DeleteData();
                            PH_PY030_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            // fm_VIEW_MODE
                            PH_PY030_LoadCaption();
                            PH_PY030_MTX01();
                        }
                        else
                        {
                        }

                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY030_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "BtnPrint2")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY030_Print_Report02);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY030")
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
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", "");  // 기본정보-사번                  //기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SMSTCOD", ""); // 조회조건-사번                   //조회조건-사번
                }
                else if (pVal.BeforeAction == false)
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if ((pVal.ItemUID == "Mat01"))
                        {
                        }
                        else
                        {
                            PH_PY030_FlushToItemValue(pVal.ItemUID);
                            if (pVal.ItemUID == "MSTCOD")
                            {
                                oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'", "");                                //성명
                            }
                            else if (pVal.ItemUID == "SMSTCOD")
                            {
                                oForm.Items.Item("SMSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.VALUE + "'", "");                                //성명
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY030_FlushToItemValue(pVal.ItemUID);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY030_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    PH_PY030_FormResize();
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// ROW_DELETE(Raise_FormMenuEvent에서 호출)
        /// 해당 클래스에서는 사용되지 않음
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
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
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PH_PY030A.RemoveRecord(oDS_PH_PY030A.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY030_Add_MatrixRow(0);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY030A.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY030_Add_MatrixRow(oMat01.RowCount);
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

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oForm.Freeze(true);

                            oDS_PH_PY030A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE);                  // 관리번호
                            oDS_PH_PY030A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.VALUE);                    // 사업장
                            oDS_PH_PY030A.SetValue("U_DestNo1", 0, oMat01.Columns.Item("DestNo1").Cells.Item(pVal.Row).Specific.VALUE);                  // 출장번호1
                            oDS_PH_PY030A.SetValue("U_DestNo2", 0, oMat01.Columns.Item("DestNo2").Cells.Item(pVal.Row).Specific.VALUE);                  // 출장번호2
                            oDS_PH_PY030A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE);                    // 사원번호
                            oDS_PH_PY030A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.VALUE);                    // 사원성명
                            oDS_PH_PY030A.SetValue("U_Destinat", 0, oMat01.Columns.Item("Destinat").Cells.Item(pVal.Row).Specific.VALUE);                // 출장지
                            oDS_PH_PY030A.SetValue("U_Dest2", 0, oMat01.Columns.Item("Dest2").Cells.Item(pVal.Row).Specific.VALUE);                      // 출장지상세
                            oDS_PH_PY030A.SetValue("U_CoCode", 0, oMat01.Columns.Item("CoCode").Cells.Item(pVal.Row).Specific.VALUE);                    // 작번
                            oDS_PH_PY030A.SetValue("U_FrDate", 0, oMat01.Columns.Item("FrDate").Cells.Item(pVal.Row).Specific.VALUE.Replace(".", ""));   // 시작일자
                            oDS_PH_PY030A.SetValue("U_FrTime", 0, oMat01.Columns.Item("FrTime").Cells.Item(pVal.Row).Specific.VALUE);                    // 시작시각
                            oDS_PH_PY030A.SetValue("U_ToDate", 0, oMat01.Columns.Item("ToDate").Cells.Item(pVal.Row).Specific.VALUE.Replace(".", ""));   // 종료일자
                            oDS_PH_PY030A.SetValue("U_ToTime", 0, oMat01.Columns.Item("ToTime").Cells.Item(pVal.Row).Specific.VALUE);                    // 종료시각
                            oDS_PH_PY030A.SetValue("U_Object", 0, oMat01.Columns.Item("Object").Cells.Item(pVal.Row).Specific.VALUE);                    // 목적
                            oDS_PH_PY030A.SetValue("U_Comments", 0, oMat01.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.VALUE);                // 비고
                            oDS_PH_PY030A.SetValue("U_RegCls", 0, oMat01.Columns.Item("RegCls").Cells.Item(pVal.Row).Specific.VALUE);                    // 등록구분
                            oDS_PH_PY030A.SetValue("U_ObjCls", 0, oMat01.Columns.Item("ObjCls").Cells.Item(pVal.Row).Specific.VALUE);                    // 목적구분
                            oDS_PH_PY030A.SetValue("U_DestCode", 0, oMat01.Columns.Item("DestCode").Cells.Item(pVal.Row).Specific.VALUE);                // 출장지역
                            oDS_PH_PY030A.SetValue("U_DestDiv", 0, oMat01.Columns.Item("DestDiv").Cells.Item(pVal.Row).Specific.VALUE);                  // 출장구분
                            oDS_PH_PY030A.SetValue("U_Vehicle", 0, oMat01.Columns.Item("Vehicle").Cells.Item(pVal.Row).Specific.VALUE);                  // 차량구분
                            oDS_PH_PY030A.SetValue("U_FuelPrc", 0, oMat01.Columns.Item("FuelPrc").Cells.Item(pVal.Row).Specific.VALUE);                  // 1L단가
                            oDS_PH_PY030A.SetValue("U_FuelType", 0, oMat01.Columns.Item("FuelType").Cells.Item(pVal.Row).Specific.VALUE);                // 유류
                            oDS_PH_PY030A.SetValue("U_Distance", 0, oMat01.Columns.Item("Distance").Cells.Item(pVal.Row).Specific.VALUE);                // 거리
                            oDS_PH_PY030A.SetValue("U_TransExp", 0, oMat01.Columns.Item("TransExp").Cells.Item(pVal.Row).Specific.VALUE);                // 교통비
                            oDS_PH_PY030A.SetValue("U_DayExp", 0, oMat01.Columns.Item("DayExp").Cells.Item(pVal.Row).Specific.VALUE);                    // 일비
                            oDS_PH_PY030A.SetValue("U_FoodNum", 0, oMat01.Columns.Item("FoodNum").Cells.Item(pVal.Row).Specific.VALUE);                  // 식수
                            oDS_PH_PY030A.SetValue("U_FoodExp", 0, oMat01.Columns.Item("FoodExp").Cells.Item(pVal.Row).Specific.VALUE);                  // 식비
                            oDS_PH_PY030A.SetValue("U_ParkExp", 0, oMat01.Columns.Item("ParkExp").Cells.Item(pVal.Row).Specific.VALUE);                  // 주차비
                            oDS_PH_PY030A.SetValue("U_TollExp", 0, oMat01.Columns.Item("TollExp").Cells.Item(pVal.Row).Specific.VALUE);                  // 도로비
                            oDS_PH_PY030A.SetValue("U_TotalExp", 0, oMat01.Columns.Item("TotalExp").Cells.Item(pVal.Row).Specific.VALUE);                // 합계

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PH_PY030_LoadCaption();
                            oForm.Freeze(false);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY030A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY030B);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":                            // 취소
                            break;
                        case "1286":                            // 닫기
                            break;
                        case "1293":                            // 행삭제
                            break;
                        case "1281":                            // 찾기
                            break;
                        case "1282":                            // 추가
                            // 추가버튼 클릭시 메트릭스 insertrow
                            PH_PY030_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PH_PY030_LoadCaption();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":                            // 레코드이동버튼
                            break;
                        case "7169":
                            // 엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            oForm.Freeze(true);
                            PH_PY030_Add_MatrixRow(oMat01.VisualRowCount);
                            oForm.Freeze(false);
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":                            // 취소
                            break;
                        case "1286":                            // 닫기
                            break;
                        case "1293":                            // 행삭제
                            break;
                        case "1281":                            // 찾기
                            break;
                        case "1282":                            // 추가
                            break;

                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "7169":
                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PH_PY030B.RemoveRecord(oDS_PH_PY030B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
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
            string sQry = string.Empty;

            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
                            break;
                    }
                    ////BeforeAction = False
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
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

        /// <summary>
        /// PH_PY030_Print_Report01 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY030_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;
            string CLTCOD = string.Empty;
            string DestNo1 = string.Empty;
            string DestNo2 = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD  = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                DestNo1 = oForm.Items.Item("DestNo1").Specific.VALUE.ToString().Trim();
                DestNo2 = oForm.Items.Item("DestNo2").Specific.VALUE.ToString().Trim();

                WinTitle = "[PH_PY030] 공용증";

                if (CLTCOD == "1")//창원
                {
                    ReportName = "PH_PY030_01.rpt";

                }
                else if (CLTCOD == "2")//동래
                {
                    ReportName = "PH_PY030_02.rpt";

                }
                else if (CLTCOD == "3")//사상
                {
                    ReportName = "PH_PY030_03.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
               // List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                //dataPackFormula.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); 

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo1", DestNo1));
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo2", DestNo2));


                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                // formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY030_Print_Report02 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY030_Print_Report02()
        {
            string DocNum = string.Empty;
            string WinTitle = string.Empty;
            string ReportName = string.Empty;
            string sQry = string.Empty;
            string DocEntry = string.Empty;            // 관리번호
            string CLTCOD = string.Empty;              // 사업장
            string DestNo1 = string.Empty;             // 출장번호1
            string DestNo2 = string.Empty;             // 출장번호2
            string MSTCOD = string.Empty;              // 사번
            string Destinat = string.Empty;            // 출장지(%)
            string Dest2 = string.Empty;               // 출장지상세
            string CoCode = string.Empty;              // 작번(%)
            //string FrDt = string.Empty;                // 기간(Fr)
            //string ToDt = string.Empty;                // 기간(To)
            System.DateTime FrDt = default(System.DateTime);
            System.DateTime ToDt = default(System.DateTime);
            string Object_Renamed = string.Empty;      // 목적(%)
            string Comments = string.Empty;            // 비고(%)
            string TeamCode = string.Empty;            // 팀
            string RegCls = string.Empty;              // 등록구분
            string ObjCls = string.Empty;              // 목적구분
            string DestCode = string.Empty;            // 출장지역
            string DestDiv = string.Empty;             // 출장구분
            string Vehicle = string.Empty;             // 차량구분

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("SDocEntry").Specific.VALUE;        
                CLTCOD   = oForm.Items.Item("SCLTCOD").Specific.Selected.VALUE.ToString().Trim(); 
                DestNo1  = oForm.Items.Item("SDestNo1").Specific.VALUE.ToString().Trim();        
                DestNo2  = oForm.Items.Item("SDestNo2").Specific.VALUE.ToString().Trim();        
                MSTCOD   = oForm.Items.Item("SMSTCOD").Specific.VALUE.ToString().Trim();          
                Destinat = oForm.Items.Item("SDestinat").Specific.Selected.VALUE.ToString().Trim();    
                Dest2    = oForm.Items.Item("SDest2").Specific.VALUE.ToString().Trim();
                CoCode   = oForm.Items.Item("SCoCode").Specific.VALUE.ToString().Trim();
                //FrDt     = oForm.Items.Item("SFrDate").Specific.VALUE.ToString().Trim();
                //ToDt     = oForm.Items.Item("SToDate").Specific.VALUE.ToString().Trim();
                FrDt = DateTime.ParseExact(oForm.Items.Item("SFrDate").Specific.Value, "yyyyMMdd", null);
                ToDt = DateTime.ParseExact(oForm.Items.Item("SToDate").Specific.Value, "yyyyMMdd", null);
                Object_Renamed = oForm.Items.Item("SObject").Specific.VALUE.ToString().Trim();
                Comments = oForm.Items.Item("SComments").Specific.VALUE.ToString().Trim();
                TeamCode = oForm.Items.Item("STeamCode").Specific.VALUE.ToString().Trim();
                RegCls   = oForm.Items.Item("SRegCls").Specific.Selected.VALUE.ToString().Trim();
                ObjCls   = oForm.Items.Item("SObjCls").Specific.Selected.VALUE.ToString().Trim();
                DestCode = oForm.Items.Item("SDestCode").Specific.Selected.VALUE.ToString().Trim();
                DestDiv  = oForm.Items.Item("SDestDiv").Specific.Selected.VALUE.ToString().Trim();
                Vehicle  = oForm.Items.Item("SVehicle").Specific.Selected.VALUE.ToString().Trim();

                WinTitle = "[PH_PY030] 공용여비 신청서";

                if (CLTCOD == "1")//창원
                {
                    ReportName = "PH_PY030_04.rpt";

                }
                else if (CLTCOD == "2")//동래
                {
                    ReportName = "PH_PY030_05.rpt";

                }
                else if (CLTCOD == "3")//사상
                {
                    ReportName = "PH_PY030_06.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                // List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                //dataPackFormula.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); 

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); 
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD));
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo1", DestNo1));
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo2", DestNo2));
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD));
                dataPackParameter.Add(new PSH_DataPackClass("@Destinat", Destinat));
                dataPackParameter.Add(new PSH_DataPackClass("@Dest2", Dest2));
                dataPackParameter.Add(new PSH_DataPackClass("@CoCode", CoCode));
                dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
                dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
                dataPackParameter.Add(new PSH_DataPackClass("@Object", Object_Renamed));
                dataPackParameter.Add(new PSH_DataPackClass("@Comments", Comments));
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                dataPackParameter.Add(new PSH_DataPackClass("@RegCls", RegCls));
                dataPackParameter.Add(new PSH_DataPackClass("@ObjCls", ObjCls));
                dataPackParameter.Add(new PSH_DataPackClass("@DestCode", DestCode));
                dataPackParameter.Add(new PSH_DataPackClass("@DestDiv", DestDiv));
                dataPackParameter.Add(new PSH_DataPackClass("@Vehicle", Vehicle));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                // formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_Print_Report02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
