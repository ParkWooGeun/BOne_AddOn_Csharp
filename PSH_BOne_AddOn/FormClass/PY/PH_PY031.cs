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
    /// 출장등록
    /// </summary>
    internal class PH_PY031 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY031A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY031B;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY031.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY031_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY031");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code";
                
                oForm.Freeze(true);
                PH_PY031_CreateItems();
                PH_PY031_ComboBox_Setting();
                PH_PY031_EnableMenus();
                PH_PY031_SetDocument(oFormDocEntry01);
                PH_PY031_FormResize();
                PH_PY031_LoadCaption();
                PH_PY031_FormItemEnabled();
                PH_PY031_FormReset(); //폼초기화 추가(2013.01.29 송명규)
                
                // 기간
                oForm.Items.Item("SFrDate").Specific.Value = DateTime.Now.ToString("yyyy.MM");
                oForm.Items.Item("SToDate").Specific.Value = DateTime.Now.ToString("yyyy.MM");
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
        private void PH_PY031_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY031A = oForm.DataSources.DBDataSources.Item("@PH_PY031A");
                oDS_PH_PY031B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

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
                oForm.DataSources.UserDataSources.Add("SDest1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDest1").Specific.DataBind.SetBound(true, "", "SDest1");

                // 행선지
                oForm.DataSources.UserDataSources.Add("SDest2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SDest2").Specific.DataBind.SetBound(true, "", "SDest2");

                // 작번
                oForm.DataSources.UserDataSources.Add("SCoCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SCoCode").Specific.DataBind.SetBound(true, "", "SCoCode");

                // 시작월
                oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

                // 종료월
                oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
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

                // 출장구분
                oForm.DataSources.UserDataSources.Add("SDestDiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestDiv").Specific.DataBind.SetBound(true, "", "SDestDiv");

                // 출장지역
                oForm.DataSources.UserDataSources.Add("SDestCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDestCode").Specific.DataBind.SetBound(true, "", "SDestCode");

                // 팀
                oForm.DataSources.UserDataSources.Add("STeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STeamCode").Specific.DataBind.SetBound(true, "", "STeamCode");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY031_LoadCaption()
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_LoadCaption_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY031_Add_MatrixRow
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY031_Add_MatrixRow(int oRow, bool RowIserted = false)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PH_PY031B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY031B.Offset = oRow;
                oDS_PH_PY031B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY031_MTX01
        /// </summary>
        private void PH_PY031_MTX01()
        {
            int i;
            int ErrNum = 0;
            string sQry;
            string sDocEntry = string.Empty;            // 관리번호
            string sCLTCOD;              // 사업장
            string SDestNo1;             // 출장번호1
            string SDestNo2;             // 출장번호2
            string sMSTCOD;              // 사원번호
            string SDest1;               // 출장지
            string SDest2;               // 행선지
            string SCoCode;              // 작번
            string SFrDate;              // 시작일자
            string SToDate;              // 종료일자
            string SObject;              // 목적
            string SComments;            // 비고
            string SRegCls;              // 등록구분
            string SDestDiv;             // 출장구분
            string SDestCode;            // 출장지역
            string sTeamCode;            // 팀

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                sCLTCOD   = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                SDestNo1  = oForm.Items.Item("SDestNo1").Specific.Value.ToString().Trim();
                SDestNo2  = oForm.Items.Item("SDestNo2").Specific.Value.ToString().Trim();
                sMSTCOD   = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim();
                SDest1    = oForm.Items.Item("SDest1").Specific.Value.ToString().Trim();
                SDest2    = oForm.Items.Item("SDest2").Specific.Value.ToString().Trim();
                SCoCode   = oForm.Items.Item("SCoCode").Specific.Value.ToString().Trim();
                SFrDate   = oForm.Items.Item("SFrDate").Specific.Value.ToString().Trim().Replace(".", "");
                SToDate   = oForm.Items.Item("SToDate").Specific.Value.ToString().Trim().Replace(".", "");
                SObject   = oForm.Items.Item("SObject").Specific.Value.ToString().Trim();
                SComments = oForm.Items.Item("SComments").Specific.Value.ToString().Trim();
                SRegCls   = oForm.Items.Item("SRegCls").Specific.Value.ToString().Trim();
                SDestDiv  = oForm.Items.Item("SDestDiv").Specific.Value.ToString().Trim();
                SDestCode = oForm.Items.Item("SDestCode").Specific.Value.ToString().Trim();
                sTeamCode = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim();

                sQry = "EXEC [PH_PY031_01] '";
                sQry += sDocEntry + "','";              // 관리번호
                sQry += sCLTCOD + "','";                // 사업장
                sQry += SDestNo1 + "','";               // 출장번호1
                sQry += SDestNo2 + "','";               // 출장번호2
                sQry += sMSTCOD + "','";                // 사원번호
                sQry += SDest1 + "','";                 // 출장지
                sQry += SDest2 + "','";                 // 행선지
                sQry += SCoCode + "','";                // 작번
                sQry += SFrDate + "','";                // 시작일자
                sQry += SToDate + "','";                // 종료일자
                sQry += SObject + "','";                // 목적
                sQry += SComments + "','";              // 비고
                sQry += SRegCls + "','";                // 등록구분
                sQry += SDestDiv + "','";               // 출장구분
                sQry += SDestCode + "','";              // 출장지역
                sQry += sTeamCode + "'";               // 팀
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY031B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY031_LoadCaption();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY031B.Size)
                    {
                        oDS_PH_PY031B.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PH_PY031B.Offset = i;

                    oDS_PH_PY031B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY031B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());  // 관리번호
                    oDS_PH_PY031B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim());    // 사업장
                    oDS_PH_PY031B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("DestNo1").Value.ToString().Trim());   // 출장번호1
                    oDS_PH_PY031B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim());   // 출장번호2
                    oDS_PH_PY031B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());    // 사원번호
                    oDS_PH_PY031B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("MSTNAM").Value.ToString().Trim());    // 사원성명
                    oDS_PH_PY031B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("Dest1").Value.ToString().Trim());     // 출장지
                    oDS_PH_PY031B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("Dest2").Value.ToString().Trim());     // 행선지
                    oDS_PH_PY031B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("CoCode").Value.ToString().Trim());    // 작번
                    oDS_PH_PY031B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("FrDate").Value.ToString().Trim());    // 시작일자
                    oDS_PH_PY031B.SetValue("U_ColTm01", i,  oRecordSet01.Fields.Item("FrTime").Value.ToString().Trim());    // 시작시각
                    oDS_PH_PY031B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("ToDate").Value.ToString().Trim());    // 종료일자
                    oDS_PH_PY031B.SetValue("U_ColTm02", i,  oRecordSet01.Fields.Item("ToTime").Value.ToString().Trim());    // 종료시각
                    oDS_PH_PY031B.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("Object").Value.ToString().Trim());    // 목적
                    oDS_PH_PY031B.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim());  // 비고
                    oDS_PH_PY031B.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("RegCls").Value.ToString().Trim());    // 등록구분
                    oDS_PH_PY031B.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("DestDiv").Value.ToString().Trim());   // 출장구분
                    oDS_PH_PY031B.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("DestCode").Value.ToString().Trim());  // 출장지역
                    oDS_PH_PY031B.SetValue("U_ColPtg01", i, oRecordSet01.Fields.Item("PayPstg").Value.ToString().Trim());   // 지급율
                    oDS_PH_PY031B.SetValue("U_ColRat01", i, oRecordSet01.Fields.Item("Rate").Value.ToString().Trim());      // 환율
                    oDS_PH_PY031B.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("TransExp").Value.ToString().Trim());  // 교통비
                    oDS_PH_PY031B.SetValue("U_ColSum02", i, oRecordSet01.Fields.Item("InsurExp").Value.ToString().Trim());  // 보험료
                    oDS_PH_PY031B.SetValue("U_ColSum03", i, oRecordSet01.Fields.Item("AirpExp").Value.ToString().Trim());   // 공항세
                    oDS_PH_PY031B.SetValue("U_ColReg24", i, oRecordSet01.Fields.Item("FoodNum").Value.ToString().Trim());   // 식수
                    oDS_PH_PY031B.SetValue("U_ColSum04", i, oRecordSet01.Fields.Item("FoodExp").Value.ToString().Trim());   // 식비
                    oDS_PH_PY031B.SetValue("U_ColSum05", i, oRecordSet01.Fields.Item("LodgExp").Value.ToString().Trim());   // 숙박비
                    oDS_PH_PY031B.SetValue("U_ColSum06", i, oRecordSet01.Fields.Item("FLodgExp").Value.ToString().Trim());  // 숙박비(외화)
                    oDS_PH_PY031B.SetValue("U_ColSum07", i, oRecordSet01.Fields.Item("DayExp").Value.ToString().Trim());    // 일비
                    oDS_PH_PY031B.SetValue("U_ColSum08", i, oRecordSet01.Fields.Item("FDayExp").Value.ToString().Trim());   // 일비(외화)
                    oDS_PH_PY031B.SetValue("U_ColSum09", i, oRecordSet01.Fields.Item("TotalExp").Value.ToString().Trim());  // 합계

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
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// PH_PY031_DeleteData
        /// </summary>
        private void PH_PY031_DeleteData()
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

                    sQry = "SELECT COUNT(*) FROM [@PH_PY031A] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.RecordCount == 0)
                    {
                        ErrNum = 1;
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "EXEC PH_PY031_04 '" + DocEntry + "'";
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
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_DeleteData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY031_UpdateData
        /// </summary>
        /// <returns></returns>
        private bool PH_PY031_UpdateData()
        {
            bool functionReturnValue = false;
            string sQry;
            int DocEntry;                      // 관리번호
            string CLTCOD;            // 사업장
            string DestNo1;           // 출장번호1
            string DestNo2;           // 출장번호2
            string MSTCOD;            // 사원번호
            string MSTNAM;            // 사원성명
            string Dest1;             // 출장지
            string Dest2;             // 행선지
            string CoCode;            // 작번
            string FrDate;            // 시작일자
            string FrTime;            // 시작시각
            string ToDate;            // 종료일자
            string ToTime;            // 종료시각
            string purpose;    // 목적
            string Comments;          // 비고
            string RegCls;            // 등록구분
            string DestDiv;           // 출장구분
            string DestCode;          // 출장지역
            double PayPstg;                      // 지급율
            double currencyRate;                 // 환율
            double TransExp;                     // 교통비
            double InsurExp;                     // 보험료
            double AirpExp;                      // 공항세
            string FoodNum;           // 식수
            double FoodExp;                      // 식비
            double LodgExp;                      // 숙박비
            double FLodgExp;                     // 숙박비(외화)
            double DayExp;                       // 일비
            double FDayExp;                      // 일비(외화)
            double TotalExp;                     // 합계

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
                Dest1    = oForm.Items.Item("Dest1").Specific.Value.ToString().Trim();
                Dest2    = oForm.Items.Item("Dest2").Specific.Value.ToString().Trim();
                CoCode   = oForm.Items.Item("CoCode").Specific.Value.ToString().Trim();
                FrDate   = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
                FrTime   = oForm.Items.Item("FrTime").Specific.Value.ToString().Trim();
                ToDate   = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
                ToTime   = oForm.Items.Item("ToTime").Specific.Value.ToString().Trim();
                purpose = oForm.Items.Item("Object").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.Value.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.Value.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.Value.ToString().Trim();
                PayPstg  = Convert.ToDouble(oForm.Items.Item("PayPstg").Specific.Value.ToString().Trim());
                currencyRate = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                InsurExp = Convert.ToDouble(oForm.Items.Item("InsurExp").Specific.Value.ToString().Trim());
                AirpExp  = Convert.ToDouble(oForm.Items.Item("AirpExp").Specific.Value.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.Value.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                LodgExp  = Convert.ToDouble(oForm.Items.Item("LodgExp").Specific.Value.ToString().Trim());
                FLodgExp = Convert.ToDouble(oForm.Items.Item("FLodgExp").Specific.Value.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                FDayExp  = Convert.ToDouble(oForm.Items.Item("FDayExp").Specific.Value.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.Value.ToString().Trim());

                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).ToString().Trim()))
                {
                    dataHelpClass.MDC_GF_Message("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E");
                    throw new Exception();
                }

                sQry = "EXEC [PH_PY031_03] '";
                sQry += DocEntry + "','";              // 관리번호
                sQry += CLTCOD + "','";                // 사업장
                sQry += DestNo1 + "','";               // 출장번호1
                sQry += DestNo2 + "','";               // 출장번호2
                sQry += MSTCOD + "','";                // 사원번호
                sQry += MSTNAM + "','";                // 사원성명
                sQry += Dest1 + "','";                 // 출장지
                sQry += Dest2 + "','";                 // 행선지
                sQry += CoCode + "','";                // 작번
                sQry += FrDate + "','";                // 시작일자
                sQry += FrTime + "','";                // 시작시각
                sQry += ToDate + "','";                // 종료일자
                sQry += ToTime + "','";                // 종료시각
                sQry += purpose + "','";        // 목적
                sQry += Comments + "','";              // 비고
                sQry += RegCls + "','";                // 등록구분
                sQry += DestDiv + "','";               // 출장구분
                sQry += DestCode + "','";              // 출장지역
                sQry += PayPstg + "','";               // 지급율
                sQry += currencyRate + "','";          // 환율
                sQry += TransExp + "','";              // 교통비
                sQry += InsurExp + "','";              // 보험료
                sQry += AirpExp + "','";               // 공항세
                sQry += FoodNum + "','";               // 식수
                sQry += FoodExp + "','";               // 식비
                sQry += LodgExp + "','";               // 숙박비
                sQry += FLodgExp + "','";              // 숙박비(외화)
                sQry += DayExp + "','";                // 일비
                sQry += FDayExp + "','";               // 일비(외화)
                sQry += TotalExp + "'";               // 합계

                oRecordSet01.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("수정 완료!", "S");

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_UpdateData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY031_DeleteData
        /// </summary>
        private bool PH_PY031_AddData()
        {
            bool functionReturnValue = false;
            string sQry;
            int DocEntry;                      // 관리번호
            string CLTCOD;            // 사업장
            string DestNo1;           // 출장번호1
            string DestNo2;           // 출장번호2
            string MSTCOD;            // 사원번호
            string MSTNAM;            // 사원성명
            string Dest1;             // 출장지
            string Dest2;             // 행선지
            string CoCode;            // 작번
            string FrDate;            // 시작일자
            string FrTime;            // 시작시각
            string ToDate;            // 종료일자
            string ToTime;            // 종료시각
            string purpose;    // 목적
            string Comments;          // 비고
            string RegCls;            // 등록구분
            string DestDiv;           // 출장구분
            string DestCode;          // 출장지역
            double PayPstg;                      // 지급율
            double currencyRate;                 // 환율
            double TransExp;                     // 교통비
            double InsurExp;                     // 보험료
            double AirpExp;                      // 공항세
            string FoodNum;           // 식수
            double FoodExp;                      // 식비
            double LodgExp;                      // 숙박비
            double FLodgExp;                     // 숙박비(외화)
            double DayExp;                       // 일비
            double FDayExp;                      // 일비(외화)
            double TotalExp;                     // 합계
            string UserSign;          // UserSign

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                CLTCOD  = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DestNo1 = oForm.Items.Item("DestNo1").Specific.Value.ToString().Trim();
                DestNo2 = oForm.Items.Item("DestNo2").Specific.Value.ToString().Trim();
                MSTCOD  = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                MSTNAM  = oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim();
                Dest1   = oForm.Items.Item("Dest1").Specific.Value.ToString().Trim();
                Dest2   = oForm.Items.Item("Dest2").Specific.Value.ToString().Trim();
                CoCode  = oForm.Items.Item("CoCode").Specific.Value.ToString().Trim();
                FrDate  = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
                FrTime  = oForm.Items.Item("FrTime").Specific.Value.ToString().Trim();
                ToDate  = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
                ToTime  = oForm.Items.Item("ToTime").Specific.Value.ToString().Trim();
                purpose = oForm.Items.Item("Object").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                RegCls   = oForm.Items.Item("RegCls").Specific.Value.ToString().Trim();
                DestDiv  = oForm.Items.Item("DestDiv").Specific.Value.ToString().Trim();
                DestCode = oForm.Items.Item("DestCode").Specific.Value.ToString().Trim();
                PayPstg = Convert.ToDouble(oForm.Items.Item("PayPstg").Specific.Value.ToString().Trim());
                currencyRate = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                InsurExp = Convert.ToDouble(oForm.Items.Item("InsurExp").Specific.Value.ToString().Trim());
                AirpExp  = Convert.ToDouble(oForm.Items.Item("AirpExp").Specific.Value.ToString().Trim());
                FoodNum  = oForm.Items.Item("FoodNum").Specific.Value.ToString().Trim();
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                LodgExp  = Convert.ToDouble(oForm.Items.Item("LodgExp").Specific.Value.ToString().Trim());
                FLodgExp = Convert.ToDouble(oForm.Items.Item("FLodgExp").Specific.Value.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                FDayExp  = Convert.ToDouble(oForm.Items.Item("FDayExp").Specific.Value.ToString().Trim());
                TotalExp = Convert.ToDouble(oForm.Items.Item("TotalExp").Specific.Value.ToString().Trim());
                UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PH_PY031A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }

                sQry = "EXEC [PH_PY031_02] '";
                sQry += DocEntry + "','";              // 관리번호
                sQry += CLTCOD + "','";                // 사업장
                sQry += DestNo1 + "','";               // 출장번호1
                sQry += DestNo2 + "','";               // 출장번호2
                sQry += MSTCOD + "','";                // 사원번호
                sQry += MSTNAM + "','";                // 사원성명
                sQry += Dest1 + "','";                 // 출장지
                sQry += Dest2 + "','";                 // 행선지
                sQry += CoCode + "','";                // 작번
                sQry += FrDate + "','";                // 시작일자
                sQry += FrTime + "','";                // 시작시각
                sQry += ToDate + "','";                // 종료일자
                sQry += ToTime + "','";                // 종료시각
                sQry += purpose + "','";        // 목적
                sQry += Comments + "','";              // 비고
                sQry += RegCls + "','";                // 등록구분
                sQry += DestDiv + "','";               // 출장구분
                sQry += DestCode + "','";              // 출장지역
                sQry += PayPstg + "','";               // 지급율
                sQry += currencyRate + "','";          // 환율
                sQry += TransExp + "','";              // 교통비
                sQry += InsurExp + "','";              // 보험료
                sQry += AirpExp + "','";               // 공항세
                sQry += FoodNum + "','";               // 식수
                sQry += FoodExp + "','";               // 식비
                sQry += LodgExp + "','";               // 숙박비
                sQry += FLodgExp + "','";              // 숙박비(외화)
                sQry += DayExp + "','";                // 일비
                sQry += FDayExp + "','";               // 일비(외화)
                sQry += TotalExp + "','";              // 합계
                sQry += UserSign + "'";               // UserSign

                oRecordSet02.DoQuery(sQry);
                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_AddData_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private bool PH_PY031_HeaderSpaceLineDel()
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
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.Value.ToString().Trim()))  // 종료일자
                {
                    ErrNum = 6;
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
                else if (ErrNum == 6)
                {
                    dataHelpClass.MDC_GF_Message("종료일자는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        private void PH_PY031_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
        {
            int loopCount;
            string sQry;
            double currencyRate;      // 환율
            double LodgExp;           // 숙박비
            double FLodgExp;          // 숙박비(외화)
            double DayExp;            // 일비
            double FDayExp;           // 일비(외화)

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {

                    case "TransExp":                        // 교통비
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "InsurExp":                        // 보험료
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "AirpExp":                        // 공항세
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "FoodNum":                        // 식수
                        PH_PY031_CalculateFoodExp();       // 식비 계산
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "FoodExp":                        // 식비
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "LodgExp":                        // 숙박비
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "DayExp":                        // 일비
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "FrDate":                        // 시작일자
                        PH_PY031_GetDestNo();
                        break;
                    case "CLTCOD":                        // 사업장
                        PH_PY031_GetDestNo();              // 출장번호생성
                        break;
                    case "DestDiv":                        // 출장구분
                        //출장지역 Reset
                        oDS_PH_PY031A.SetValue("U_DestCode", 0, "%");                        // 금액 필드들을 Reset
                        oDS_PH_PY031A.SetValue("U_PayPstg", 0, Convert.ToString(100));       // 지급율 100으로 재설정
                        oDS_PH_PY031A.SetValue("U_Rate", 0, Convert.ToString(1));            // 환율 1로 재설정
                        oDS_PH_PY031A.SetValue("U_TransExp", 0, "0");        // 교통비
                        oDS_PH_PY031A.SetValue("U_InsurExp", 0, "0");        // 보험료
                        oDS_PH_PY031A.SetValue("U_AirpExp", 0, "0");         // 공항세
                        oDS_PH_PY031A.SetValue("U_FoodNum", 0, "0");         // 식수
                        oDS_PH_PY031A.SetValue("U_FoodExp", 0, "0");         // 식비
                        oDS_PH_PY031A.SetValue("U_LodgExp", 0, "0");         // 숙박비
                        oDS_PH_PY031A.SetValue("U_FLodgExp", 0, "0");        // 숙박비(외화)
                        oDS_PH_PY031A.SetValue("U_DayExp", 0, "0");          // 일비
                        oDS_PH_PY031A.SetValue("U_FDayExp", 0, "0");         // 일비(외화)
                        oDS_PH_PY031A.SetValue("U_TotalExp", 0, "0");        // 합계

                        if (oDS_PH_PY031A.GetValue("U_DestDiv", 0).ToString().Trim() == "2")
                        {
                            oForm.Items.Item("Rate").Enabled = true;
                            // 해외출장인 경우는 환율 Enable = True
                            oForm.Items.Item("FoodNum").Enabled = false;
                            // 해외출장인 경우는 식수 Enable = False
                            oForm.Items.Item("FLodgExp").Enabled = true;
                            // 해외출장인 경우는 숙박비(외화) Enable = True
                            oForm.Items.Item("FDayExp").Enabled = true;
                            // 해외줄장인 경우는 일비(외화) Enable = True
                            oForm.Items.Item("LodgExp").Enabled = false;
                            // 해외출장인 경우는 숙박비 Enable = False
                            oForm.Items.Item("DayExp").Enabled = false;
                            // 해외줄장인 경우는 일비 Enable = False
                        }
                        else
                        {
                            oForm.Items.Item("Rate").Enabled = false;
                            // 이외인 경우는 환율 Enable = False
                            oForm.Items.Item("FoodNum").Enabled = true;
                            // 이외인 경우는 식수 Enable = True
                            oForm.Items.Item("FLodgExp").Enabled = false;
                            // 이외인 경우는 숙박비(외화) Enable = False
                            oForm.Items.Item("FDayExp").Enabled = false;
                            // 이외인 경우는 일비(외화) Enable = False
                            oForm.Items.Item("LodgExp").Enabled = true;
                            // 이외인 경우는 숙박비 Enable = True
                            oForm.Items.Item("DayExp").Enabled = true;
                            // 이외인 경우는 일비 Enable = True
                        }
                        break;
                    case "DestCode":                        // 출장지역
                        PH_PY031_GetLodgExpAndDayExp();                        // 숙박비, 일비 조회
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "PayPstg":                        // 지급율
                        PH_PY031_CalculateTotalExp();
                        break;
                    case "Rate":
                    case "FLodgExp":
                    case "FDayExp":                        // 환율, 숙박비(외화), 일비(외화)
                        currencyRate = Convert.ToDouble(oDS_PH_PY031A.GetValue("U_Rate", 0).ToString().Trim());
                        FLodgExp = Convert.ToDouble(oDS_PH_PY031A.GetValue("U_FLodgExp", 0).ToString().Trim());
                        FDayExp = Convert.ToDouble(oDS_PH_PY031A.GetValue("U_FDayExp", 0).ToString().Trim());

                        LodgExp = ((FLodgExp * currencyRate) / 10) * 10;                      // 원단위 절사
                        DayExp = ((FDayExp * currencyRate) / 10) * 10;                        // 원단위 절사

                        oDS_PH_PY031A.SetValue("U_LodgExp", 0, Convert.ToString(LodgExp));
                        oDS_PH_PY031A.SetValue("U_DayExp", 0, Convert.ToString(DayExp));

                        PH_PY031_CalculateTotalExp();
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
                        sQry += "               AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                        sQry += "               AND U_UseYN = 'Y'";
                        sQry += " ORDER BY  U_Seq";
                        dataHelpClass.Set_ComboList((oForm.Items.Item("STeamCode").Specific), sQry, "", false, false);
                        oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        oForm.Items.Item("STeamCode").DisplayDesc = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY031_ComboBox_Setting
        /// </summary>
        private void PH_PY031_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // 기본정보
                // 등록구분
                oForm.Items.Item("RegCls").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P223'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("RegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("RegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("DestDiv").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P235'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("DestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("DestCode").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P225'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("DestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("DestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 식수
                //    Call oForm.Items("FoodNum").Specific.ValidValues.Add("0", "선택")
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P227'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("FoodNum").Specific, sQry, "", false, false);
                oForm.Items.Item("FoodNum").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 조회정보
                // 등록구분
                oForm.Items.Item("SRegCls").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P223'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SRegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("SRegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장구분
                oForm.Items.Item("SDestDiv").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P235'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestDiv").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 출장지역
                oForm.Items.Item("SDestCode").Specific.ValidValues.Add("%", "전체");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P225'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("SDestCode").Specific, sQry, "", false, false);
                oForm.Items.Item("SDestCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 매트릭스
                // 사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");

                // 등록구분
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P223'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RegCls"), sQry, "", "");

                //출장구분
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P235'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestDiv"), sQry, "", "");

                // 출장지역
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P225'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("DestCode"), sQry, "", "");

                // 식수
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P227'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("FoodNum"), sQry, "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_ComboBox_Setting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY031_EnableMenus()
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY031_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY031_FormItemEnabled();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY031_FormResize
        /// </summary>
        private void PH_PY031_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_FormResize_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY031_FormReset
        /// </summary>
        private void PH_PY031_FormReset()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                //관리번호
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PH_PY031A]";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PH_PY031A.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PH_PY031A.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1));
                }

                string User_BPLID = null;
                User_BPLID = dataHelpClass.User_BPLID();

                // 기준정보
                oDS_PH_PY031A.SetValue("U_CLTCOD", 0, User_BPLID);        // 사업장
                oDS_PH_PY031A.SetValue("U_DestNo1", 0, "");               // 출장번호1
                oDS_PH_PY031A.SetValue("U_DestNo2", 0, "");               // 출장번호2
                oDS_PH_PY031A.SetValue("U_MSTCOD", 0, "");                // 사원번호
                oDS_PH_PY031A.SetValue("U_MSTNAM", 0, "");                // 사원성명
                oDS_PH_PY031A.SetValue("U_Dest1", 0, "");                 // 출장지
                oDS_PH_PY031A.SetValue("U_Dest2", 0, "");                 // 행선지
                oDS_PH_PY031A.SetValue("U_CoCode", 0, "");                // 작번
                oDS_PH_PY031A.SetValue("U_FrDate", 0, DateTime.Now.ToString("yyyyMMdd"));                 // 시작일자
                oDS_PH_PY031A.SetValue("U_FrTime", 0, "");                // 시작시각
                oDS_PH_PY031A.SetValue("U_ToDate", 0, DateTime.Now.ToString("yyyyMMdd"));                 // 종료일자
                oDS_PH_PY031A.SetValue("U_ToTime", 0, "");                // 종료시각
                oDS_PH_PY031A.SetValue("U_Object", 0, "");                // 목적
                oDS_PH_PY031A.SetValue("U_Comments", 0, "");              // 비고
                oDS_PH_PY031A.SetValue("U_RegCls", 0, "01");              // 등록구분
                oDS_PH_PY031A.SetValue("U_DestDiv", 0, "%");              // 출장구분
                oDS_PH_PY031A.SetValue("U_DestCode", 0, "%");             // 출장지역
                oDS_PH_PY031A.SetValue("U_PayPstg", 0, "100");           // 지급율
                oDS_PH_PY031A.SetValue("U_Rate", 0, "1");                // 환율
                oDS_PH_PY031A.SetValue("U_TransExp", 0, "0");            // 교통비
                oDS_PH_PY031A.SetValue("U_InsurExp", 0, "0");            // 보험료
                oDS_PH_PY031A.SetValue("U_AirpExp", 0, "0");             // 공항세
                oDS_PH_PY031A.SetValue("U_FoodNum", 0, "0");                             // 식수
                oDS_PH_PY031A.SetValue("U_FoodExp", 0, "0");             // 식비
                oDS_PH_PY031A.SetValue("U_LodgExp", 0, "0");             // 숙박비
                oDS_PH_PY031A.SetValue("U_LodgExp", 0, "0");             // 숙박비(외화)
                oDS_PH_PY031A.SetValue("U_DayExp", 0, "0");              // 일비
                oDS_PH_PY031A.SetValue("U_FDayExp", 0, "0");             // 일비(외화)
                oDS_PH_PY031A.SetValue("U_TotalExp", 0, "0");            // 합계
                //출장번호
                PH_PY031_GetDestNo();

                // Control 설정
                oForm.Items.Item("Rate").Enabled = false;
                // 기본적으로 환율은 Enable = False
                oForm.Items.Item("FLodgExp").Enabled = false;
                // 기본적으로 숙박비(외화)은 Enable = False
                oForm.Items.Item("FDayExp").Enabled = false;
                // 기본적으로 일비(외화)는 Enable = False

                oForm.Items.Item("MSTCOD").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_FormReset_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY031_CalculateTransExp
        /// </summary>
        private void PH_PY031_CalculateTransExp()
        {
            double FuelPrc;                //유류단가
            double Distance;                //거리
            double TransExp;                //교통비

            try
            {
                FuelPrc = Convert.ToDouble(oForm.Items.Item("FuelPrc").Specific.Value.ToString().Trim());
                Distance = Convert.ToDouble(oForm.Items.Item("Distance").Specific.Value.ToString().Trim());
                TransExp = ((FuelPrc * Distance * 0.1) / 10) * 10;                //원단위 절사

                oForm.Items.Item("TransExp").Specific.Value = TransExp;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_CalculateTransExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY031_CalculateTransExp
        /// </summary>
        private void PH_PY031_CalculateTotalExp()
        {
            double TransExp;          // 교통비
            double InsurExp;          // 보험료
            double AirpExp;           // 공항세
            double FoodExp;           // 식비
            double LodgExp;           // 숙박비
            double DayExp;            // 일비
            double TotalExp;          // 합계
            double PayPstg;           // 지급율

            try
            {
                TransExp = Convert.ToDouble(oForm.Items.Item("TransExp").Specific.Value.ToString().Trim());
                InsurExp = Convert.ToDouble(oForm.Items.Item("InsurExp").Specific.Value.ToString().Trim());
                AirpExp  = Convert.ToDouble(oForm.Items.Item("AirpExp").Specific.Value.ToString().Trim());
                FoodExp  = Convert.ToDouble(oForm.Items.Item("FoodExp").Specific.Value.ToString().Trim());
                LodgExp  = Convert.ToDouble(oForm.Items.Item("LodgExp").Specific.Value.ToString().Trim());
                DayExp   = Convert.ToDouble(oForm.Items.Item("DayExp").Specific.Value.ToString().Trim());
                PayPstg  = Convert.ToDouble(oForm.Items.Item("PayPstg").Specific.Value.ToString().Trim());
                TotalExp = (TransExp + InsurExp + AirpExp + FoodExp + LodgExp + DayExp) * (PayPstg / 100);
                oDS_PH_PY031A.SetValue("U_TotalExp", 0, Convert.ToString(TotalExp));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_CalculateTotalExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY031_GetDestNo
        /// </summary>
        private void PH_PY031_GetDestNo()
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

                sQry = "EXEC PH_PY031_05 '" + CLTCOD + "', '" + FrDate + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("DestNo1").Specific.Value = FrDate;
                oForm.Items.Item("DestNo2").Specific.Value = oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_GetDestNo_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY031_GetLodgExpAndDayExp
        /// </summary>
        private void PH_PY031_GetLodgExpAndDayExp()
        {
            short ErrNum = 0;
            string MSTCOD;                // 사번
            string DestCode;              // 출장지역
            string CodeString = string.Empty;            // 여비코드
            string currencyRate;          // 환율
            string sQry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사번을 선택하지 않으면
                if (string.IsNullOrEmpty(oDS_PH_PY031A.GetValue("U_MSTCOD", 0).ToString().Trim()) & oDS_PH_PY031A.GetValue("U_DestCode", 0).ToString().Trim() != "%")
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                MSTCOD = oDS_PH_PY031A.GetValue("U_MSTCOD", 0).ToString().Trim();         // 사번
                DestCode = oDS_PH_PY031A.GetValue("U_DestCode", 0).ToString().Trim();     // 출장구분
                currencyRate = oDS_PH_PY031A.GetValue("U_Rate", 0).ToString().Trim();     // 환율

                //국내
                if (DestCode == "01")
                {
                    CodeString = "P232";                    // 일본
                }
                else if (DestCode == "02")
                {
                    CodeString = "P233";                    // 특지
                }
                else if (DestCode == "03")
                {
                    CodeString = "P234";                    // 갑지
                }
                else if (DestCode == "04")
                {
                    CodeString = "P230";                    // 을지
                }
                else if (DestCode == "05")
                {
                    CodeString = "P231";
                }

                sQry = "  SELECT    T1.U_Num1 AS [LodgExp],";
                sQry += "           T1.U_Num2 AS [DayExp]";
                sQry += " FROM      [@PH_PY001A] AS T0";
                sQry += "           LEFT JOIN";
                sQry += "           [@PS_HR200L] AS T1";
                sQry += "               ON T0.U_JIGCOD = T1.U_Code";
                sQry += "               AND T1.Code = '" + CodeString + "'";
                sQry += "               AND T1.U_UseYN = 'Y'";
                sQry += " WHERE     T0.Code = '" + MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);
                
                if (DestCode == "01") //출장지역 국내
                {
                    oDS_PH_PY031A.SetValue("U_LodgExp", 0, oRecordSet01.Fields.Item("LodgExp").Value.ToString().Trim());
                    oDS_PH_PY031A.SetValue("U_DayExp", 0,  oRecordSet01.Fields.Item("DayExp").Value.ToString().Trim());
                    oDS_PH_PY031A.SetValue("U_FLodgExp", 0, "0");
                    oDS_PH_PY031A.SetValue("U_FDayExp", 0, "0");
                }
                else //출장지역 국내 이외
                {
                    oDS_PH_PY031A.SetValue("U_FLodgExp", 0, oRecordSet01.Fields.Item("LodgExp").Value.ToString().Trim());
                    oDS_PH_PY031A.SetValue("U_FDayExp", 0,  oRecordSet01.Fields.Item("DayExp").Value.ToString().Trim());
                    oDS_PH_PY031A.SetValue("U_LodgExp", 0, Convert.ToString(((Convert.ToDouble(oRecordSet01.Fields.Item("LodgExp").Value.ToString().Trim()) * Convert.ToDouble(currencyRate)) / 10) * 10));
                    oDS_PH_PY031A.SetValue("U_DayExp", 0, Convert.ToString(((Convert.ToDouble(oRecordSet01.Fields.Item("DayExp").Value.ToString().Trim()) * Convert.ToDouble(currencyRate)) / 10) * 10));
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사원을 먼저 선택하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oDS_PH_PY031A.SetValue("U_DestCode", 0, "%");
                    oForm.Items.Item("MSTCOD").Click();
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_GetLodgExpAndDayExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PH_PY031_GetFuelPrc
        /// </summary>
        private void PH_PY031_GetFuelPrc()
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
                CLTCOD = oDS_PH_PY031A.GetValue("U_CLTCOD", 0).ToString().Trim();                // 사업장

                if (!string.IsNullOrEmpty(oDS_PH_PY031A.GetValue("U_FrDate", 0).ToString().Trim()))
                {
                    StdYear = oDS_PH_PY031A.GetValue("U_FrDate", 0).Substring(0, 4).ToString().Trim();
                    StdMonth = oDS_PH_PY031A.GetValue("U_FrDate", 0).Substring(4, 2).ToString().Trim();
                }

                FuelType = oDS_PH_PY031A.GetValue("U_FuelType", 0).ToString().Trim();            // 유류

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
                else if (FuelType == "3") //경유
                {
                    FuelPrice = oRecordSet01.Fields.Item("Diesel").Value;
                }
                else
                {
                    FuelPrice = 0;
                }
                oDS_PH_PY031A.SetValue("U_FuelPrc", 0, Convert.ToString(FuelPrice));
                oForm.Items.Item("Distance").Click();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_GetFuelPrc_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 식비 계산
        /// </summary>
        private void PH_PY031_CalculateFoodExp()
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
                if (string.IsNullOrEmpty(oDS_PH_PY031A.GetValue("U_MSTCOD", 0).ToString().Trim()) && oDS_PH_PY031A.GetValue("U_FoodNum", 0).ToString().Trim() != "0")
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                MSTCOD = oDS_PH_PY031A.GetValue("U_MSTCOD", 0).ToString().Trim();                       // 사번
                FoodNum = Convert.ToInt16(oDS_PH_PY031A.GetValue("U_FoodNum", 0).ToString().Trim());    // 식수

                sQry = "  SELECT    T1.U_Num4 AS [FoodPrc]";
                sQry += " FROM      [@PH_PY001A] AS T0";
                sQry += "           LEFT JOIN";
                sQry += "           [@PS_HR200L] AS T1";
                sQry += "               ON T0.U_JIGCOD = T1.U_Code";
                sQry += "               AND T1.Code = 'P232'";
                sQry += "               AND T1.U_UseYN = 'Y'";
                sQry += " WHERE     T0.Code = '" + MSTCOD + "'";

                oRecordSet01.DoQuery(sQry);

                FoodPrc = oRecordSet01.Fields.Item("FoodPrc").Value;
                FoodExp = FoodPrc * FoodNum;

                oDS_PH_PY031A.SetValue("U_FoodExp", 0, Convert.ToString(FoodExp));
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_CalculateFoodExp_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY031_FormItemEnabled()
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY031_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY031_Print_Report01 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY031_Print_Report01()
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

                WinTitle = "[PH_PY031] 여비지출전표";

                if (CLTCOD == "1") //창원
                {
                    ReportName = "PH_PY031_01.rpt";

                }
                else if (CLTCOD == "2") //동래
                {
                    ReportName = "PH_PY031_02.rpt";

                }
                else if (CLTCOD == "3") //사상
                {
                    ReportName = "PH_PY031_03.rpt";
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY031_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //20
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    // 추가/확인 버튼클릭
                    if (pVal.ItemUID == "BtnAdd")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY031_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY031_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY031_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY031_LoadCaption();
                            PH_PY031_MTX01();

                            oLast_Mode = (int)oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY031_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PH_PY031_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PH_PY031_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY031_LoadCaption();
                            PH_PY031_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") // 조회
                    {
                        PH_PY031_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        PH_PY031_LoadCaption();
                        PH_PY031_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") // 삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            PH_PY031_DeleteData();
                            PH_PY031_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY031_LoadCaption();
                            PH_PY031_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY031_Print_Report01);
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
                            PH_PY031_FlushToItemValue(pVal.ItemUID);
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
                    PH_PY031_FlushToItemValue(pVal.ItemUID);
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
                    PH_PY031_FormItemEnabled();
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
                    PH_PY031_FormResize();
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
                if ((oLastColRow01 > 0))
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
                        oDS_PH_PY031A.RemoveRecord(oDS_PH_PY031A.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY031_Add_MatrixRow(0);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY031A.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY031_Add_MatrixRow(oMat01.RowCount);
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

                            //DataSource를 이용하여 각 컨트롤에 값을 출력
                            oDS_PH_PY031A.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);                          // 관리번호
                            oDS_PH_PY031A.SetValue("U_CLTCOD", 0, oMat01.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.Value);                            // 사업장
                            oDS_PH_PY031A.SetValue("U_DestNo1", 0, oMat01.Columns.Item("DestNo1").Cells.Item(pVal.Row).Specific.Value);                          // 출장번호1
                            oDS_PH_PY031A.SetValue("U_DestNo2", 0, oMat01.Columns.Item("DestNo2").Cells.Item(pVal.Row).Specific.Value);                          // 출장번호2
                            oDS_PH_PY031A.SetValue("U_MSTCOD", 0, oMat01.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value);                            // 사원번호
                            oDS_PH_PY031A.SetValue("U_MSTNAM", 0, oMat01.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value);                            // 사원성명
                            oDS_PH_PY031A.SetValue("U_Dest1", 0, oMat01.Columns.Item("Dest1").Cells.Item(pVal.Row).Specific.Value);                              // 출장지
                            oDS_PH_PY031A.SetValue("U_Dest2", 0, oMat01.Columns.Item("Dest2").Cells.Item(pVal.Row).Specific.Value);                              // 행선지
                            oDS_PH_PY031A.SetValue("U_CoCode", 0, oMat01.Columns.Item("CoCode").Cells.Item(pVal.Row).Specific.Value);                            // 작번
                            oDS_PH_PY031A.SetValue("U_FrDate", 0, oMat01.Columns.Item("FrDate").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));           // 시작일자
                            oDS_PH_PY031A.SetValue("U_FrTime", 0, oMat01.Columns.Item("FrTime").Cells.Item(pVal.Row).Specific.Value);                            // 시작시각
                            oDS_PH_PY031A.SetValue("U_ToDate", 0, oMat01.Columns.Item("ToDate").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));           // 종료일자
                            oDS_PH_PY031A.SetValue("U_ToTime", 0, oMat01.Columns.Item("ToTime").Cells.Item(pVal.Row).Specific.Value);                            // 종료시각
                            oDS_PH_PY031A.SetValue("U_Object", 0, oMat01.Columns.Item("Object").Cells.Item(pVal.Row).Specific.Value);                            // 목적
                            oDS_PH_PY031A.SetValue("U_Comments", 0, oMat01.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.Value);                        // 비고
                            oDS_PH_PY031A.SetValue("U_RegCls", 0, oMat01.Columns.Item("RegCls").Cells.Item(pVal.Row).Specific.Value);                            // 등록구분
                            oDS_PH_PY031A.SetValue("U_DestDiv", 0, oMat01.Columns.Item("DestDiv").Cells.Item(pVal.Row).Specific.Value);                          // 출장구분
                            oDS_PH_PY031A.SetValue("U_DestCode", 0, oMat01.Columns.Item("DestCode").Cells.Item(pVal.Row).Specific.Value);                        // 출장지역
                            oDS_PH_PY031A.SetValue("U_PayPstg", 0, oMat01.Columns.Item("PayPstg").Cells.Item(pVal.Row).Specific.Value);                          // 지급율
                            oDS_PH_PY031A.SetValue("U_Rate", 0, oMat01.Columns.Item("Rate").Cells.Item(pVal.Row).Specific.Value);                                // 환율
                            oDS_PH_PY031A.SetValue("U_TransExp", 0, oMat01.Columns.Item("TransExp").Cells.Item(pVal.Row).Specific.Value);                        // 교통비
                            oDS_PH_PY031A.SetValue("U_InsurExp", 0, oMat01.Columns.Item("InsurExp").Cells.Item(pVal.Row).Specific.Value);                        // 보험료
                            oDS_PH_PY031A.SetValue("U_AirpExp", 0, oMat01.Columns.Item("AirpExp").Cells.Item(pVal.Row).Specific.Value);                          // 공항세
                            oDS_PH_PY031A.SetValue("U_FoodNum", 0, oMat01.Columns.Item("FoodNum").Cells.Item(pVal.Row).Specific.Value);                          // 식수
                            oDS_PH_PY031A.SetValue("U_FoodExp", 0, oMat01.Columns.Item("FoodExp").Cells.Item(pVal.Row).Specific.Value);                          // 식비
                            oDS_PH_PY031A.SetValue("U_LodgExp", 0, oMat01.Columns.Item("LodgExp").Cells.Item(pVal.Row).Specific.Value);                          // 숙박비
                            oDS_PH_PY031A.SetValue("U_FLodgExp", 0, oMat01.Columns.Item("FLodgExp").Cells.Item(pVal.Row).Specific.Value);                        // 숙박비(외화)
                            oDS_PH_PY031A.SetValue("U_DayExp", 0, oMat01.Columns.Item("DayExp").Cells.Item(pVal.Row).Specific.Value);                            // 일비
                            oDS_PH_PY031A.SetValue("U_FDayExp", 0, oMat01.Columns.Item("FDayExp").Cells.Item(pVal.Row).Specific.Value);                          // 일비(외화)
                            oDS_PH_PY031A.SetValue("U_TotalExp", 0, oMat01.Columns.Item("TotalExp").Cells.Item(pVal.Row).Specific.Value);                        // 합계
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PH_PY031_LoadCaption();

                            //숙박비,일비 컨드롤 및 숙박비(외화), 일비(외화) 컨트롤 Enable 설정
                            //해외출장
                            if (oMat01.Columns.Item("DestDiv").Cells.Item(pVal.Row).Specific.Value == "2")
                            {
                                oForm.Items.Item("Rate").Enabled = true;
                                oForm.Items.Item("FLodgExp").Enabled = true;
                                oForm.Items.Item("FDayExp").Enabled = true;
                                oForm.Items.Item("LodgExp").Enabled = false;
                                oForm.Items.Item("DayExp").Enabled = false;
                            }
                            else
                            {
                                oForm.Items.Item("Rate").Enabled = false;
                                oForm.Items.Item("FLodgExp").Enabled = false;
                                oForm.Items.Item("FDayExp").Enabled = false;
                                oForm.Items.Item("LodgExp").Enabled = true;
                                oForm.Items.Item("DayExp").Enabled = true;
                            }

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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY031A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY031B);
                }
                else if (pVal.Before_Action == false)
                {   
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
                            PH_PY031_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PH_PY031_LoadCaption();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":                            // 레코드이동버튼
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
