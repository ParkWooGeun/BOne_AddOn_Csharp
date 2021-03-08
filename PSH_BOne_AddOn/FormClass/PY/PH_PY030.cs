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
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY030A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY030B;
        private string oLastItemUID01; // 클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01;  // 마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01;     // 마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Mode;        // 마지막 모드
        
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

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY030_CreateItems();
                PH_PY030_ComboBox_Setting();
                PH_PY030_EnableMenus();
                PH_PY030_SetDocument(oFormDocEntry01);
                PH_PY030_FormResize();
                PH_PY030_LoadCaption();
                PH_PY030_FormItemEnabled();
                PH_PY030_FormReset();
                
                oMat01.Columns.Item("Check").Visible = false; // 선택 체크박스 Visible = False
                
                // 기간
                oForm.Items.Item("SFrDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                oForm.Items.Item("SToDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                // 사번 포커스
                oForm.Items.Item("MSTCOD").Click();
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
        private void PH_PY030_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PH_PY030B.InsertRecord(oRow);
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
            int i;
            int ErrNum = 0;
            string sQry;
            string sDocEntry;    // 관리번호
            string sCLTCOD;      // 사업장
            string SDestNo1;     // 출장번호1
            string SDestNo2;     // 출장번호2
            string sMSTCOD;      // 사원번호
            string SDestinat;    // 출장지
            string SDest2;       // 출장지상세
            string SCoCode;      // 작번
            string SFrDate;      // 시작일자
            string SToDate;      // 종료일자
            string SObject;      // 목적
            string SComments;    // 비고
            string SRegCls;      // 등록구분
            string SObjCls;      // 목적구분
            string SDestCode;    // 출장지역
            string SDestDiv;     // 출장구분
            string SVehicle;     // 차량구분
            string sTeamCode;    // 팀(2013.06.01 송명규 추가)

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                sDocEntry = oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim();                // 관리번호
                sCLTCOD   = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();                  // 사업장
                SDestNo1  = oForm.Items.Item("SDestNo1").Specific.Value.ToString().Trim();                 // 출장번호1
                SDestNo2  = oForm.Items.Item("SDestNo2").Specific.Value.ToString().Trim();                 // 출장번호2
                sMSTCOD   = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim();                  // 사원번호
                SDestinat = oForm.Items.Item("SDestinat").Specific.Value.ToString().Trim();                // 출장지
                SDest2    = oForm.Items.Item("SDest2").Specific.Value.ToString().Trim();                   // 출장지상세
                SCoCode   = oForm.Items.Item("SCoCode").Specific.Value.ToString().Trim();                  // 작번
                SFrDate   = oForm.Items.Item("SFrDate").Specific.Value.ToString().Trim().Replace(".", ""); // 시작일자
                SToDate   = oForm.Items.Item("SToDate").Specific.Value.ToString().Trim().Replace(".", ""); // 종료일자
                SObject   = oForm.Items.Item("SObject").Specific.Value.ToString().Trim();                  // 목적
                SComments = oForm.Items.Item("SComments").Specific.Value.ToString().Trim();                // 비고
                SRegCls   = oForm.Items.Item("SRegCls").Specific.Value.ToString().Trim();                  // 등록구분
                SObjCls   = oForm.Items.Item("SObjCls").Specific.Value.ToString().Trim();                  // 목적구분
                SDestCode = oForm.Items.Item("SDestCode").Specific.Value.ToString().Trim();                // 출장지역
                SDestDiv  = oForm.Items.Item("SDestDiv").Specific.Value.ToString().Trim();                 // 출장구분
                SVehicle  = oForm.Items.Item("SVehicle").Specific.Value.ToString().Trim();                 // 차량구분
                sTeamCode = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim();                // 팀

                sQry = "EXEC [PH_PY030_01] '";
                sQry += sDocEntry + "','";
                sQry += sCLTCOD + "','";
                sQry += SDestNo1 + "','";
                sQry += SDestNo2 + "','";
                sQry += sMSTCOD + "','";
                sQry += SDestinat + "','";
                sQry += SDest2 + "','";
                sQry += SCoCode + "','";
                sQry += SFrDate + "','";
                sQry += SToDate + "','";
                sQry += SObject + "','";
                sQry += SComments + "','";
                sQry += SRegCls + "','";
                sQry += SObjCls + "','";
                sQry += SDestCode + "','";
                sQry += SDestDiv + "','";
                sQry += SVehicle + "','";
                sQry += sTeamCode + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY030B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
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
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
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
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                
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
            string sQry;
            string DocEntry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {

                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PH_PY030A] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.RecordCount == 0)
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

            string sQry;
            int DocEntry; // 관리번호
            string CLTCOD; // 사업장
            string DestNo1; // 출장번호1
            string DestNo2; // 출장번호2
            string MSTCOD; // 사원번호
            string MSTNAM; // 사원성명
            string Destinat; // 출장지
            string Dest2; // 출장지상세
            string CoCode; // 작번
            string FrDate; // 시작일자
            string FrTime; // 시작시각
            string ToDate; // 종료일자
            string ToTime; // 종료시각
            string purpose; // 목적
            string Comments; // 비고
            string RegCls; // 등록구분
            string ObjCls; // 목적구분
            string DestCode; // 출장지역
            string DestDiv; // 출장구분
            string Vehicle; // 차량구분
            double FuelPrc; // 1L단가
            string FuelType; // 유류
            double Distance; // 거리
            double TransExp; // 교통비
            double DayExp; // 일비
            string FoodNum; // 식수
            double FoodExp; // 식비
            double ParkExp; // 주차비
            double TollExp; // 도로비
            double TotalExp; // 합계

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
                CLTCOD   = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DestNo1  = oForm.Items.Item("DestNo1").Specific.Value.ToString().Trim();
                DestNo2  = oForm.Items.Item("DestNo2").Specific.Value.ToString().Trim();
                MSTCOD   = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                MSTNAM   = oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim();
                Destinat = oForm.Items.Item("Destinat").Specific.Value.ToString().Trim();
                Dest2    = oForm.Items.Item("Dest2").Specific.Value.ToString().Trim();
                CoCode   = oForm.Items.Item("CoCode").Specific.Value.ToString().Trim();
                FrDate   = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
                FrTime   = oForm.Items.Item("FrTime").Specific.Value.ToString().Trim();
                ToDate   = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
                ToTime   = oForm.Items.Item("ToTime").Specific.Value.ToString().Trim();
                purpose = oForm.Items.Item("Object").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.Value.ToString().Trim();
                ObjCls   = oForm.Items.Item("ObjCls").Specific.Value.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.Value.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.Value.ToString().Trim();
                Vehicle  = oForm.Items.Item("Vehicle").Specific.Value.ToString().Trim();
                FuelPrc  = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.Value.ToString().Trim());
                FuelType = oForm.Items.Item("FuelType").Specific.Value.ToString().Trim();
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.Value.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.Value.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                ParkExp  = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.Value.ToString().Trim());
                TollExp  = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.Value.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.Value.ToString().Trim());


                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).Trim()))
                {
                    dataHelpClass.MDC_GF_Message("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E");
                    functionReturnValue = false;
                    throw new Exception();
                }

                sQry = "EXEC [PH_PY030_03] '";
                sQry += DocEntry + "','";              // 관리번호
                sQry += CLTCOD + "','";                // 사업장
                sQry += DestNo1 + "','";               // 출장번호1
                sQry += DestNo2 + "','";               // 출장번호2
                sQry += MSTCOD + "','";                // 사원번호
                sQry += MSTNAM + "','";                // 사원성명
                sQry += Destinat + "','";              // 출장지
                sQry += Dest2 + "','";                 // 출장지상세
                sQry += CoCode + "','";                // 작번
                sQry += FrDate + "','";                // 시작일자
                sQry += FrTime + "','";                // 시작시각
                sQry += ToDate + "','";                // 종료일자
                sQry += ToTime + "','";                // 종료시각
                sQry += purpose + "','";        // 목적
                sQry += Comments + "','";              // 비고
                sQry += RegCls + "','";                // 등록구분
                sQry += ObjCls + "','";                // 목적구분
                sQry += DestCode + "','";              // 출장지역
                sQry += DestDiv + "','";               // 출장구분
                sQry += Vehicle + "','";               // 차량구분
                sQry += FuelPrc + "','";               // 1L단가
                sQry += FuelType + "','";              // 유류
                sQry += Distance + "','";              // 거리
                sQry += TransExp + "','";              // 교통비
                sQry += DayExp + "','";                // 일비
                sQry += FoodNum + "','";               // 식수
                sQry += FoodExp + "','";               // 식비
                sQry += ParkExp + "','";               // 주차비
                sQry += TollExp + "','";               // 도로비
                sQry += TotalExp + "'";               // 합계

                oRecordSet01.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("수정 완료!", "S");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY030_UpdateData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY030_DeleteData
        /// </summary>
        private bool PH_PY030_AddData()
        {
            bool functionReturnValue = false;
            int DocEntry;
            string sQry;
            string CLTCOD; // 사업장
            string DestNo1; // 출장번호1
            string DestNo2; // 출장번호2
            string MSTCOD; // 사원번호
            string MSTNAM; // 사원성명
            string Destinat; // 출장지
            string Dest2; // 출장지상세
            string CoCode; // 작번
            string FrDate; // 시작일자
            string FrTime; // 시작시각
            string ToDate; // 종료일자
            string ToTime; // 종료시각
            string purpose; // 목적
            string Comments; // 비고
            string RegCls; // 등록구분
            string ObjCls; // 목적구분
            string DestCode; // 출장지역
            string DestDiv; // 출장구분
            string Vehicle; // 차량구분
            double FuelPrc; // 1L단가
            string FuelType; // 유류
            double Distance; // 거리
            double TransExp; // 교통비
            double DayExp; // 일비
            string FoodNum; // 식수
            double FoodExp; // 식비
            double ParkExp; // 주차비
            double TollExp; // 도로비
            double TotalExp; // 합계
            string UserSign; // UserSign

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                CLTCOD   = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DestNo1  = oForm.Items.Item("DestNo1").Specific.Value.ToString().Trim();
                DestNo2  = oForm.Items.Item("DestNo2").Specific.Value.ToString().Trim();
                MSTCOD   = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                MSTNAM   = oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim();
                Destinat = oForm.Items.Item("Destinat").Specific.Value.ToString().Trim();
                Dest2    = oForm.Items.Item("Dest2").Specific.Value.ToString().Trim();
                CoCode   = oForm.Items.Item("CoCode").Specific.Value.ToString().Trim();
                FrDate   = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
                FrTime   = oForm.Items.Item("FrTime").Specific.Value.ToString().Trim();
                ToDate   = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
                ToTime   = oForm.Items.Item("ToTime").Specific.Value.ToString().Trim();
                purpose = oForm.Items.Item("Object").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.Value.ToString().Trim();
                ObjCls   = oForm.Items.Item("ObjCls").Specific.Value.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.Value.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.Value.ToString().Trim();
                Vehicle  = oForm.Items.Item("Vehicle").Specific.Value.ToString().Trim();
                FuelPrc  = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.Value.ToString().Trim());
                FuelType = oForm.Items.Item("FuelType").Specific.Value.ToString().Trim();
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.Value.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.Value.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                ParkExp  = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.Value.ToString().Trim());
                TollExp  = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.Value.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.Value.ToString().Trim());
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

                sQry = "EXEC [PH_PY030_02] '";
                sQry += DocEntry + "','";       // 관리번호
                sQry += CLTCOD + "','";         // 사업장
                sQry += DestNo1 + "','";        // 출장번호1
                sQry += DestNo2 + "','";        // 출장번호2
                sQry += MSTCOD + "','";         // 사원번호
                sQry += MSTNAM + "','";         // 사원성명
                sQry += Destinat + "','";       // 출장지
                sQry += Dest2 + "','";          // 출장지상세
                sQry += CoCode + "','";         // 작번
                sQry += FrDate + "','";         // 시작일자
                sQry += FrTime + "','";         // 시작시각
                sQry += ToDate + "','";         // 종료일자
                sQry += ToTime + "','";         // 종료시각
                sQry += purpose + "','"; // 목적
                sQry += Comments + "','";       // 비고
                sQry += RegCls + "','";         // 등록구분
                sQry += ObjCls + "','";         // 목적구분
                sQry += DestCode + "','";       // 출장지역
                sQry += DestDiv + "','";        // 출장구분
                sQry += Vehicle + "','";        // 차량구분
                sQry += FuelPrc + "','";        // 1L단가
                sQry += FuelType + "','";       // 유류
                sQry += Distance + "','";       // 거리
                sQry += TransExp + "','";       // 교통비
                sQry += DayExp + "','";         // 일비
                sQry += FoodNum + "','";        // 식수
                sQry += FoodExp + "','";        // 식비
                sQry += ParkExp + "','";        // 주차비
                sQry += TollExp + "','";        // 도로비
                sQry += TotalExp + "','";       // 합계
                sQry += UserSign + "'";        // UserSign

                oRecordSet02.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
                functionReturnValue = true;
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
                if (string.IsNullOrEmpty(oForm.Items.Item("DestNo1").Specific.Value.ToString().Trim()))      // 출장번호1
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DestNo2").Specific.Value.ToString().Trim()))  // 출장번호2
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))  // 사원번호
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim()))  // 시작일자
                {
                    ErrNum = 4;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FrTime").Specific.Value.ToString().Trim()))  // 시작시각
                {
                    ErrNum = 5;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.Value.ToString().Trim()))  // 종료일자
                {
                    ErrNum = 6;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToTime").Specific.Value.ToString().Trim()))  // 종료시각
                {
                    ErrNum = 7;
                    throw new Exception();
                }

                functionReturnValue = true;
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
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY030_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int loopCount;
            string sQry;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                        sQry = "  SELECT      U_Code,";
                        sQry += "             U_CodeNm";
                        sQry += " FROM        [@PS_HR200L]";
                        sQry += " WHERE       Code = '1'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList((oForm.Items.Item("STeamCode").Specific), sQry,"", false, false);
                        oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        oForm.Items.Item("STeamCode").DisplayDesc = true;
                        break;
                    case "Destinat":
                        if (oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() == "1")
                        {
                            sQry = "Select Distance = U_Num2 From [@PS_HR200L] Where Code = 'P217' And U_Code = '" + oForm.Items.Item("Destinat").Specific.Value.ToString().Trim() + "' and U_Char1 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("Distance").Specific.Value = oRecordSet01.Fields.Item("Distance").Value;
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
        private void PH_PY030_ComboBox_Setting()
        {
            string sQry;
            //SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // 기본정보
                // 출장지
                oForm.Items.Item("Destinat").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P217'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Destinat").Specific, sQry,  "", false,  false);
                oForm.Items.Item("Destinat").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 등록구분
                oForm.Items.Item("RegCls").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P223'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("RegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("RegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 목적구분
                oForm.Items.Item("ObjCls").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P224'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("ObjCls").Specific, sQry, "", false, false);
                oForm.Items.Item("ObjCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("DestCode").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P225'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("DestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("DestDiv").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P216'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("DestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 차량구분
                oForm.Items.Item("Vehicle").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P218'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Vehicle").Specific, sQry, "", false, false);
                oForm.Items.Item("Vehicle").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 유류
                oForm.Items.Item("FuelType").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P226'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("FuelType").Specific, sQry, "", false, false);
                oForm.Items.Item("FuelType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 식수
                //    Call oForm.Items("FoodNum").Specific.ValidValues.Add("0", "선택")
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P227'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("FoodNum").Specific, sQry, "", false, false);
                oForm.Items.Item("FoodNum").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 조회정보
                // 출장지
                oForm.Items.Item("SDestinat").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P217'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestinat").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestinat").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 등록구분
                oForm.Items.Item("SRegCls").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P223'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SRegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("SRegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 목적구분
                oForm.Items.Item("SObjCls").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P224'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SObjCls").Specific, sQry, "", false, false);
                oForm.Items.Item("SObjCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("SDestCode").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P225'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("SDestDiv").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P216'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 차량구분
                oForm.Items.Item("SVehicle").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P218'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SVehicle").Specific, sQry, "", false, false);
                oForm.Items.Item("SVehicle").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 매트릭스
                // 사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId","","");

                // 출장지
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P217'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Destinat"), sQry, "", "");

                // 등록구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P223'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RegCls"), sQry, "", "");

                // 목적구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P224'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ObjCls"), sQry, "", "");

                // 출장지역
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P225'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestCode"), sQry, "", "");

                // 출장구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P216'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestDiv"), sQry, "", "");

                // 차량구분
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P218'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Vehicle"), sQry, "", "");

                // 유류
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P226'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("FuelType"), sQry, "", "");

                // 식수
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P227'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
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
                oForm.EnableMenu("1283", false);                // 삭제
                oForm.EnableMenu("1286", false);                // 닫기
                oForm.EnableMenu("1287", false);                // 복제
                oForm.EnableMenu("1285", false);                // 복원
                oForm.EnableMenu("1284", false);                // 취소
                oForm.EnableMenu("1293", false);                // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
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
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
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
        private void PH_PY030_FormReset()
        {
            string sQry;
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
                    oDS_PH_PY030A.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PH_PY030A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1));
                }

                // 기준정보
                oDS_PH_PY030A.SetValue("U_CLTCOD", 0, dataHelpClass.User_BPLID()); // 사업장
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
                oDS_PH_PY030A.SetValue("U_FuelPrc", 0, "0");                // 1L단가
                oDS_PH_PY030A.SetValue("U_FuelType", 0, "%");                               // 유류
                oDS_PH_PY030A.SetValue("U_Distance", 0, "0");               // 거리
                oDS_PH_PY030A.SetValue("U_TransExp", 0, "0");               // 교통비
                oDS_PH_PY030A.SetValue("U_DayExp", 0, "0");                 // 일비
                oDS_PH_PY030A.SetValue("U_FoodNum", 0, "0");                                // 식수
                oDS_PH_PY030A.SetValue("U_FoodExp", 0, "0");                // 식비
                oDS_PH_PY030A.SetValue("U_ParkExp", 0, "0");                // 주차비
                oDS_PH_PY030A.SetValue("U_TollExp", 0, "0");                // 도로비
                oDS_PH_PY030A.SetValue("U_TotalExp", 0, "0");               // 합계
                PH_PY030_GetDestNo(); // 출장번호

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
            double FuelPrc; //유류단가
            double Distance; //거리
            int TransExp; //교통비

            try
            {
                FuelPrc = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.Value.ToString().Trim());
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.Value.ToString().Trim());
                TransExp = Convert.ToInt32((FuelPrc * Distance * 0.1) / 10) * 10; //원단위  절사 2020.06.03 황영수

                oForm.Items.Item("TransExp").Specific.Value = TransExp;
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
            double TransExp; // 교통비
            double DayExp; // 일비
            double FoodExp; // 식비
            double ParkExp; // 주차비
            double TollExp; // 도로비
            double TotalExp; // 합계

            try
            {
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                DayExp = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                FoodExp = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                ParkExp = Convert.ToDouble(oForm.Items.Item("ParkExp").Specific.Value.ToString().Trim());
                TollExp = Convert.ToDouble(oForm.Items.Item("TollExp").Specific.Value.ToString().Trim());
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
            string FrDate;
            string CLTCOD;
            string sQry;

            PSH_CodeHelpClass codeaHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                FrDate = codeaHelpClass.Left(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim(), 6);

                sQry = "EXEC PH_PY030_05 '" + CLTCOD + "', '" + FrDate + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("DestNo1").Specific.Value = FrDate;
                oForm.Items.Item("DestNo2").Specific.Value = oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim();
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
            string CLTCOD;
            string sQry;
            string StdYear = string.Empty;
            string StdMonth = string.Empty;
            string FuelType;
            double FuelPrice;

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

                sQry = "  SELECT    T0.U_Year AS [StdYear],";
                sQry += "           T1.U_Month AS [StdMonth],";
                sQry += "           T1.U_Gasoline AS [Gasoline],";
                sQry += "           T1.U_Diesel AS [Diesel],";
                sQry += "           T1.U_LPG AS [LPG]";
                sQry += " FROM      [@PH_PY007A] AS T0";
                sQry += "           INNER JOIN";
                sQry += "           [@PH_PY007B] AS T1";
                sQry += "               ON T0.Code = T1.Code";
                sQry += " WHERE     T0.U_CLTCOD = '" + CLTCOD + "'";
                sQry += "           AND T0.U_Year = '" + StdYear + "'";
                sQry += "           AND T1.U_Month = '" + StdMonth + "'";

                oRecordSet01.DoQuery(sQry);
                
                if (FuelType == "1") //휘발유
                {
                    FuelPrice = oRecordSet01.Fields.Item("Gasoline").Value;
                }
                else if (FuelType == "2") //가스
                {
                    FuelPrice = oRecordSet01.Fields.Item("LPG").Value;
                }
                else if (FuelType == "3")  //경유
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

        /// <summary>
        /// 식비 계산
        /// </summary>
        private void PH_PY030_CalculateFoodExp()
        {
            short ErrNum = 0;
            string sQry;
            string MSTCOD;            // 사번
            short FoodNum;               // 식수
            double FoodPrc;              // 당일식비
            double FoodExp;              // 전체식비
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // 사번을 선택하지 않으면
                if (string.IsNullOrEmpty(oDS_PH_PY030A.GetValue("U_MSTCOD", 0).ToString().Trim()) && oDS_PH_PY030A.GetValue("U_FoodNum", 0).ToString().Trim() != "0")
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                MSTCOD = oDS_PH_PY030A.GetValue("U_MSTCOD", 0).ToString().Trim();                       // 사번
                FoodNum = Convert.ToInt16(oDS_PH_PY030A.GetValue("U_FoodNum", 0).ToString().Trim());    // 식수

                sQry = "   SELECT   T1.U_Num4 AS [FoodPrc]";
                sQry += "  FROM     [@PH_PY001A] AS T0";
                sQry += "           LEFT JOIN";
                sQry += "           [@PS_HR200L] AS T1";
                sQry += "               ON T0.U_JIGCOD = T1.U_Code";
                sQry += "               AND T1.Code = 'P232'";
                sQry += "               AND T1.U_UseYN = 'Y'";
                sQry += "  WHERE    T0.Code = '" + MSTCOD + "'";

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
        private void PH_PY030_FormItemEnabled()
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
        /// PH_PY030_Print_Report01 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY030_Print_Report01()
        {
            string WinTitle;
            string ReportName = string.Empty;
            string CLTCOD;
            string DestNo1;
            string DestNo2;

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DestNo1 = oForm.Items.Item("DestNo1").Specific.Value.ToString().Trim();
                DestNo2 = oForm.Items.Item("DestNo2").Specific.Value.ToString().Trim();

                WinTitle = "[PH_PY030] 공용증";

                if (CLTCOD == "1") //창원
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
                
                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo1", DestNo1));
                dataPackParameter.Add(new PSH_DataPackClass("@DestNo2", DestNo2));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
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
            string WinTitle;
            string ReportName = string.Empty;
            string DocEntry;            // 관리번호
            string CLTCOD;              // 사업장
            string DestNo1;             // 출장번호1
            string DestNo2;             // 출장번호2
            string MSTCOD;              // 사번
            string Destinat;            // 출장지(%)
            string Dest2;               // 출장지상세
            string CoCode;              // 작번(%)
            System.DateTime FrDt;
            System.DateTime ToDt;
            string purpose;      // 목적(%)
            string Comments;            // 비고(%)
            string TeamCode;            // 팀
            string RegCls;              // 등록구분
            string ObjCls;              // 목적구분
            string DestCode;            // 출장지역
            string DestDiv;             // 출장구분
            string Vehicle;             // 차량구분

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("SDocEntry").Specific.Value;
                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Selected.Value.ToString().Trim();
                DestNo1 = oForm.Items.Item("SDestNo1").Specific.Value.ToString().Trim();
                DestNo2 = oForm.Items.Item("SDestNo2").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim();
                Destinat = oForm.Items.Item("SDestinat").Specific.Selected.Value.ToString().Trim();
                Dest2 = oForm.Items.Item("SDest2").Specific.Value.ToString().Trim();
                CoCode = oForm.Items.Item("SCoCode").Specific.Value.ToString().Trim();
                FrDt = DateTime.ParseExact(oForm.Items.Item("SFrDate").Specific.Value, "yyyyMMdd", null);
                ToDt = DateTime.ParseExact(oForm.Items.Item("SToDate").Specific.Value, "yyyyMMdd", null);
                purpose = oForm.Items.Item("SObject").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("SComments").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim();
                RegCls = oForm.Items.Item("SRegCls").Specific.Selected.Value.ToString().Trim();
                ObjCls = oForm.Items.Item("SObjCls").Specific.Selected.Value.ToString().Trim();
                DestCode = oForm.Items.Item("SDestCode").Specific.Selected.Value.ToString().Trim();
                DestDiv = oForm.Items.Item("SDestDiv").Specific.Selected.Value.ToString().Trim();
                Vehicle = oForm.Items.Item("SVehicle").Specific.Selected.Value.ToString().Trim();

                WinTitle = "[PH_PY030] 공용여비 신청서";

                if (CLTCOD == "1") //창원
                {
                    ReportName = "PH_PY030_04.rpt";
                }
                else if (CLTCOD == "2") //동래
                {
                    ReportName = "PH_PY030_05.rpt";
                }
                else if (CLTCOD == "3") //사상
                {
                    ReportName = "PH_PY030_06.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                
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
                dataPackParameter.Add(new PSH_DataPackClass("@Object", purpose));
                dataPackParameter.Add(new PSH_DataPackClass("@Comments", Comments));
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                dataPackParameter.Add(new PSH_DataPackClass("@RegCls", RegCls));
                dataPackParameter.Add(new PSH_DataPackClass("@ObjCls", ObjCls));
                dataPackParameter.Add(new PSH_DataPackClass("@DestCode", DestCode));
                dataPackParameter.Add(new PSH_DataPackClass("@DestDiv", DestDiv));
                dataPackParameter.Add(new PSH_DataPackClass("@Vehicle", Vehicle));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY030_Print_Report02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    } 
                    else if (pVal.ItemUID == "BtnSearch") // 조회
                    {
                        PH_PY030_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        PH_PY030_LoadCaption();
                        PH_PY030_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") // 삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
                        {
                            PH_PY030_DeleteData();
                            PH_PY030_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", "");  // 기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SMSTCOD", ""); // 조회조건-사번
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            PH_PY030_FlushToItemValue(pVal.ItemUID, 0, "");
                            if (pVal.ItemUID == "MSTCOD")
                            {
                                oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", ""); //성명
                            }
                            else if (pVal.ItemUID == "SMSTCOD")
                            {
                                oForm.Items.Item("SMSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.Value + "'", ""); //성명
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
                    PH_PY030_FlushToItemValue(pVal.ItemUID, 0, "");
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
                        oDS_PH_PY030A.RemoveRecord(oDS_PH_PY030A.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY030_Add_MatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY030A.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY030_Add_MatrixRow(oMat01.RowCount, false);
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

                            oDS_PH_PY030A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);                  // 관리번호
                            oDS_PH_PY030A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.Value);                    // 사업장
                            oDS_PH_PY030A.SetValue("U_DestNo1", 0, oMat01.Columns.Item("DestNo1").Cells.Item(pVal.Row).Specific.Value);                  // 출장번호1
                            oDS_PH_PY030A.SetValue("U_DestNo2", 0, oMat01.Columns.Item("DestNo2").Cells.Item(pVal.Row).Specific.Value);                  // 출장번호2
                            oDS_PH_PY030A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value);                    // 사원번호
                            oDS_PH_PY030A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value);                    // 사원성명
                            oDS_PH_PY030A.SetValue("U_Destinat", 0, oMat01.Columns.Item("Destinat").Cells.Item(pVal.Row).Specific.Value);                // 출장지
                            oDS_PH_PY030A.SetValue("U_Dest2", 0, oMat01.Columns.Item("Dest2").Cells.Item(pVal.Row).Specific.Value);                      // 출장지상세
                            oDS_PH_PY030A.SetValue("U_CoCode", 0, oMat01.Columns.Item("CoCode").Cells.Item(pVal.Row).Specific.Value);                    // 작번
                            oDS_PH_PY030A.SetValue("U_FrDate", 0, oMat01.Columns.Item("FrDate").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));   // 시작일자
                            oDS_PH_PY030A.SetValue("U_FrTime", 0, oMat01.Columns.Item("FrTime").Cells.Item(pVal.Row).Specific.Value);                    // 시작시각
                            oDS_PH_PY030A.SetValue("U_ToDate", 0, oMat01.Columns.Item("ToDate").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));   // 종료일자
                            oDS_PH_PY030A.SetValue("U_ToTime", 0, oMat01.Columns.Item("ToTime").Cells.Item(pVal.Row).Specific.Value);                    // 종료시각
                            oDS_PH_PY030A.SetValue("U_Object", 0, oMat01.Columns.Item("Object").Cells.Item(pVal.Row).Specific.Value);                    // 목적
                            oDS_PH_PY030A.SetValue("U_Comments", 0, oMat01.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.Value);                // 비고
                            oDS_PH_PY030A.SetValue("U_RegCls", 0, oMat01.Columns.Item("RegCls").Cells.Item(pVal.Row).Specific.Value);                    // 등록구분
                            oDS_PH_PY030A.SetValue("U_ObjCls", 0, oMat01.Columns.Item("ObjCls").Cells.Item(pVal.Row).Specific.Value);                    // 목적구분
                            oDS_PH_PY030A.SetValue("U_DestCode", 0, oMat01.Columns.Item("DestCode").Cells.Item(pVal.Row).Specific.Value);                // 출장지역
                            oDS_PH_PY030A.SetValue("U_DestDiv", 0, oMat01.Columns.Item("DestDiv").Cells.Item(pVal.Row).Specific.Value);                  // 출장구분
                            oDS_PH_PY030A.SetValue("U_Vehicle", 0, oMat01.Columns.Item("Vehicle").Cells.Item(pVal.Row).Specific.Value);                  // 차량구분
                            oDS_PH_PY030A.SetValue("U_FuelPrc", 0, oMat01.Columns.Item("FuelPrc").Cells.Item(pVal.Row).Specific.Value);                  // 1L단가
                            oDS_PH_PY030A.SetValue("U_FuelType", 0, oMat01.Columns.Item("FuelType").Cells.Item(pVal.Row).Specific.Value);                // 유류
                            oDS_PH_PY030A.SetValue("U_Distance", 0, oMat01.Columns.Item("Distance").Cells.Item(pVal.Row).Specific.Value);                // 거리
                            oDS_PH_PY030A.SetValue("U_TransExp", 0, oMat01.Columns.Item("TransExp").Cells.Item(pVal.Row).Specific.Value);                // 교통비
                            oDS_PH_PY030A.SetValue("U_DayExp", 0, oMat01.Columns.Item("DayExp").Cells.Item(pVal.Row).Specific.Value);                    // 일비
                            oDS_PH_PY030A.SetValue("U_FoodNum", 0, oMat01.Columns.Item("FoodNum").Cells.Item(pVal.Row).Specific.Value);                  // 식수
                            oDS_PH_PY030A.SetValue("U_FoodExp", 0, oMat01.Columns.Item("FoodExp").Cells.Item(pVal.Row).Specific.Value);                  // 식비
                            oDS_PH_PY030A.SetValue("U_ParkExp", 0, oMat01.Columns.Item("ParkExp").Cells.Item(pVal.Row).Specific.Value);                  // 주차비
                            oDS_PH_PY030A.SetValue("U_TollExp", 0, oMat01.Columns.Item("TollExp").Cells.Item(pVal.Row).Specific.Value);                  // 도로비
                            oDS_PH_PY030A.SetValue("U_TotalExp", 0, oMat01.Columns.Item("TotalExp").Cells.Item(pVal.Row).Specific.Value);                // 합계

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PH_PY030_LoadCaption();
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
                oForm.Freeze(false);
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
                            PH_PY030_Add_MatrixRow(oMat01.VisualRowCount, false);
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
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
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
                else if (BusinessObjectInfo.BeforeAction == false)
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
    }
}
