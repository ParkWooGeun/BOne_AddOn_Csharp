using System;
using System.IO;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 배차신청서
    /// </summary>
    internal class PH_PY035 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY035A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY035B; //등록라인

        private string oLastItemUID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY035.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY035_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY035");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY035_CreateItems();
                PH_PY035_ComboBox_Setting();
                PH_PY035_SetDocEntry();
                PH_PY035_FormItemEnabled();
                PH_PY035_EnableMenus();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PH_PY035_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY035A = oForm.DataSources.DBDataSources.Item("@PH_PY035A");
                oDS_PH_PY035B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();


                oDS_PH_PY035A.SetValue("U_FrDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PH_PY035A.SetValue("U_ToDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }


        /// <summary>
        /// EnableMenus 메뉴설정
        /// </summary>
        private void PH_PY035_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, false, true, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY035_ComboBox_Setting
        /// </summary>
        private void PH_PY035_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                sQry = "SELECT BPLId, BPLName From[OBPL] order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("CLTCOD").Specific, "N");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //상태구분
                //oForm.Items.Item("RegCls").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT    U_Code AS [Code],";
                sQry += "           U_CodeNm As [Name]";
                sQry += " FROM      [@PS_HR200L]";
                sQry += " WHERE     Code = 'P223'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("RegCls").Specific, sQry, "", false, false);
                oForm.Items.Item("RegCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 목적구분
                //oForm.Items.Item("Object").Specific.ValidValues.Add("%", "선택");
                sQry = "  SELECT      U_Code AS [Code],";
                sQry += "             U_CodeNm As [Name]";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = 'P224'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.Set_ComboList(oForm.Items.Item("Object").Specific, sQry, "", false, false);
                oForm.Items.Item("Object").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_ComboBox_Setting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY035_SetDocEntry
        /// </summary>
        private void PH_PY035_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY035'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PH_PY035_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PH_PY035B.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PH_PY035B.Offset = oRow;
                oDS_PH_PY035B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY035_MTX01
        /// </summary>
        private bool PH_PY035_MTX01(int Count)
        {
            bool returnValue = false;
            string sQry;
            string errMessage = string.Empty;
            string sCLTCOD;              // 사업장
            string sUseCarCd;              // 사원번호
            string SFrDate;              // 시작일자
            string SToDate;              // 종료일자

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (!string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim()))
                {
                    if (Count == 1)
                    {
                        SFrDate = Convert.ToString((DateTime.ParseExact(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim(), "yyyyMMdd", null)).AddDays(-15));
                    }
                    else
                    {
                        SFrDate = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim().Replace(".", "");
                    }
                }
                else
                {
                    SFrDate = Convert.ToString(DateTime.Now.AddDays(-7));
                }

                if (!string.IsNullOrEmpty(oForm.Items.Item("UseCarCd").Specific.Value.ToString().Trim()))
                {
                    sUseCarCd = oForm.Items.Item("UseCarCd").Specific.Value.ToString().Trim();
                }
                else
                {
                    sQry = " select top(1)U_UseCarCd from [@PH_PY035A] Order BY DocEntry desc";
                    oRecordSet01.DoQuery(sQry);
                    sUseCarCd = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                }
                    
                sCLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                
                SToDate = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim().Replace(".", "");

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PH_PY035B.Clear(); //추가

                sQry = "EXEC [PH_PY035_01] '";
                sQry += sCLTCOD + "','";                // 사업장
                sQry += sUseCarCd + "','";                // 사원번호
                sQry += SFrDate + "','";                // 시작일자
                sQry += SToDate + "'";                // 종료일자
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY035B.Size)
                    {
                        oDS_PH_PY035B.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PH_PY035B.Offset = i;

                    oDS_PH_PY035B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY035B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());  // 관리번호
                    oDS_PH_PY035B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("U_RegCls").Value.ToString().Trim());   // 등록구분
                    oDS_PH_PY035B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_UseCar").Value.ToString().Trim());    // 사용차량
                    oDS_PH_PY035B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_Dest").Value.ToString().Trim());    // 목적지
                    oDS_PH_PY035B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_FrDate").Value.ToString().Trim());    // 시작일자
                    oDS_PH_PY035B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_FrTime").Value.ToString().Trim());    // 시작시간
                    oDS_PH_PY035B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_ToDate").Value.ToString().Trim());    // 종료일자
                    oDS_PH_PY035B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_ToTime").Value.ToString().Trim());    // 종료시간
                    oDS_PH_PY035B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_MSTCOD").Value.ToString().Trim());   // 신청차사번
                    oDS_PH_PY035B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("U_MSTNAM").Value.ToString().Trim());   // 신청자명
                    oDS_PH_PY035B.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("U_WMSTCOD").Value.ToString().Trim());     // 동승자사번
                    oDS_PH_PY035B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("U_WMSTNAM").Value.ToString().Trim());     // 동승자명
                    oDS_PH_PY035B.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("U_BeForKm").Value.ToString().Trim());  // 주행전Km
                    oDS_PH_PY035B.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("U_AfterKm").Value.ToString().Trim());    // 주행후Km
                    oDS_PH_PY035B.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());  // 비고
                    oRecordSet01.MoveNext();
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// report_print_035
        /// </summary>
        /// <param name="p_MSTCOD">사번</param>
        /// <param name="p_Version">문서번호</param>
        /// <returns></returns>
        private bool report_print_035(string p_Version)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PH_PY035] 배차신청서";
                ReportName = "PH_PY035_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", p_Version)); //사업장
                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("report_print_035_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY035_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("RegCls").Specific.Value.ToString().Trim()) || Convert.ToInt32(oForm.Items.Item("RegCls").Specific.Value.ToString().Trim()) != 3)
                    {
                        // 접속자에 따른 권한별 사업장 콤보박스세팅
                        dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                        oForm.EnableMenu("1281", true); //문서찾기
                        oForm.EnableMenu("1282", false); //문서추가
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CLTCOD").Enabled = true;
                        oForm.Items.Item("FrDate").Enabled = true;
                        oForm.Items.Item("ToDate").Enabled = true;
                        oForm.Items.Item("FrTime").Enabled = true;
                        oForm.Items.Item("ToTime").Enabled = true;
                        oForm.Items.Item("UseCarCd").Enabled = true;
                        oForm.Items.Item("Object").Enabled = true;
                        oForm.Items.Item("Dest").Enabled = true;
                        oForm.Items.Item("MSTCOD").Enabled = true;
                        oForm.Items.Item("WMSTCOD").Enabled = true;
                        oForm.Items.Item("Comments").Enabled = true;
                        oForm.Items.Item("RegCls").Enabled = true;
                        oForm.Items.Item("BeforKm").Enabled = false;
                        oForm.Items.Item("AfterKm").Enabled = false;
                    }
                    else 
                    {
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CLTCOD").Enabled = false;
                        oForm.Items.Item("FrDate").Enabled = false;
                        oForm.Items.Item("ToDate").Enabled = false;
                        oForm.Items.Item("FrTime").Enabled = false;
                        oForm.Items.Item("ToTime").Enabled = false;
                        oForm.Items.Item("UseCarCd").Enabled = false;
                        oForm.Items.Item("Object").Enabled = false;
                        oForm.Items.Item("Dest").Enabled = false;
                        oForm.Items.Item("MSTCOD").Enabled = false;
                        oForm.Items.Item("WMSTCOD").Enabled = false;
                        oForm.Items.Item("Comments").Enabled = false;
                        oForm.Items.Item("RegCls").Enabled = false;
                        oForm.Items.Item("BeforKm").Enabled = false;
                        oForm.Items.Item("AfterKm").Enabled = false;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    if (Convert.ToInt32(oForm.Items.Item("RegCls").Specific.Value.ToString().Trim()) == 3)
                    {
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("FrDate").Enabled = false;
                        oForm.Items.Item("ToDate").Enabled = false;
                        oForm.Items.Item("FrTime").Enabled = false;
                        oForm.Items.Item("ToTime").Enabled = false;
                        oForm.Items.Item("UseCarCd").Enabled = false;
                        oForm.Items.Item("Dest").Enabled = false;
                        oForm.Items.Item("MSTCOD").Enabled = false;
                        oForm.Items.Item("WMSTCOD").Enabled = false;
                        oForm.Items.Item("Comments").Enabled = false;
                        oForm.Items.Item("RegCls").Enabled = false;
                        oForm.Items.Item("BeforKm").Enabled = false;
                        oForm.Items.Item("AfterKm").Enabled = false;
                    }
                    else
                    {
                        // 접속자에 따른 권한별 사업장 콤보박스세팅
                        dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                        oForm.EnableMenu("1281", true); //문서찾기
                        oForm.EnableMenu("1282", true); //문서추가
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CLTCOD").Enabled = true;
                        oForm.Items.Item("FrDate").Enabled = true;
                        oForm.Items.Item("ToDate").Enabled = true;
                        oForm.Items.Item("FrTime").Enabled = true;
                        oForm.Items.Item("ToTime").Enabled = true;
                        oForm.Items.Item("UseCarCd").Enabled = true;
                        oForm.Items.Item("Object").Enabled = true;
                        oForm.Items.Item("Dest").Enabled = true;
                        oForm.Items.Item("MSTCOD").Enabled = true;
                        oForm.Items.Item("WMSTCOD").Enabled = true;
                        oForm.Items.Item("Comments").Enabled = true;
                        oForm.Items.Item("RegCls").Enabled = true;
                        oForm.Items.Item("BeforKm").Enabled = true;
                        oForm.Items.Item("AfterKm").Enabled = true;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("FrDate").Enabled = true;
                    oForm.Items.Item("ToDate").Enabled = false;
                    oForm.Items.Item("FrTime").Enabled = true;
                    oForm.Items.Item("ToTime").Enabled = false;
                    oForm.Items.Item("UseCarCd").Enabled = false;
                    oForm.Items.Item("Object").Enabled = false;
                    oForm.Items.Item("Dest").Enabled = false;
                    oForm.Items.Item("MSTCOD").Enabled = false;
                    oForm.Items.Item("WMSTCOD").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;
                    oForm.Items.Item("RegCls").Enabled = false;
                    oForm.Items.Item("BeforKm").Enabled = false;
                    oForm.Items.Item("AfterKm").Enabled = false;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY035A_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim()))  // 시작일자
                {
                    errMessage = "출발일자는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("FrTime").Specific.Value.ToString().Trim()))  // 시작일자
                {
                    errMessage = "출발시간는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.Value.ToString().Trim()))  // 종료일자
                {
                    errMessage = "도착일자는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ToTime").Specific.Value.ToString().Trim()))  // 종료일자
                {
                    errMessage = "도착시간는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Dest").Specific.Value.ToString().Trim()))  // 출장번호2
                {
                    errMessage = "목적지는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("MSTNAM").Specific.Value.ToString().Trim()))      // 출장번호1
                {
                    errMessage = "탑승자는 필수사항입니다.확인하세요";
                    throw new Exception();
                }
                returnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                    return returnValue;
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY035_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }
            return returnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
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
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = pVal.ColUID;
                        oLastColRow = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID = pVal.ItemUID;
                    oLastColUID = "";
                    oLastColRow = 0;
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                        else
                        {
                            string sVersion = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value;
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PH_PY035_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = sVersion;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
               //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
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
                    PH_PY035_FormItemEnabled();
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
        /// <summary>
        /// Raise_EVENT_KEY_DOWN
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", "");  // 기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WMSTCOD", "");  // 기본정보-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UseCarCd", ""); // 조회조건-사번
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (System.Exception ex)
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
                            if (pVal.ItemUID == "MSTCOD")
                            {
                                oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", ""); //성명
                            }
                            else if (pVal.ItemUID == "WMSTCOD")
                            {
                                oForm.Items.Item("WMSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("WMSTCOD").Specific.Value + "'", ""); //성명
                            }
                            else if (pVal.ItemUID == "UseCarCd")
                            {
                                oForm.Items.Item("UseCar").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("UseCarCd").Specific.Value + "'", ""); //차량
                                if (PH_PY035_MTX01(1) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else if (pVal.ItemUID == "FrDate" || pVal.ItemUID == "ToDate")
                            {
                                if (PH_PY035_MTX01(0) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY035A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY035B);
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            string errMessage = string.Empty;
            string sQry;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    //추가
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY035A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (!string.IsNullOrEmpty(oForm.Items.Item("BeforKm").Specific.Value.ToString().Trim()) || !string.IsNullOrEmpty(oForm.Items.Item("AfterKm").Specific.Value.ToString().Trim()))
                            {
                                sQry = " UPDATE [@PH_PY035A] SET U_RegCls ='03' WHERE DocEntry ='" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                           
                            sQry = "EXEC [PH_PY035_03] '";
                            sQry += oForm.Items.Item("UseCarCd").Specific.Value.ToString().Trim() + "','";                // 사원번호
                            sQry += oForm.Items.Item("FrDate").Specific.Value.ToString().Trim().Replace(".", "") + "','";                // 사업장
                            sQry += oForm.Items.Item("FrTime").Specific.Value.ToString().Trim()  + "','";                // 사원번호
                            sQry += oForm.Items.Item("ToDate").Specific.Value.ToString().Trim().Replace(".", "") + "','";                // 시작일자
                            sQry += oForm.Items.Item("ToTime").Specific.Value.ToString().Trim() + "'";                // 종료일자
                            oRecordSet01.DoQuery(sQry);

                            if (Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) > 0)
                            {
                                errMessage = "중복된 시간에 예약내역이 있습니다. 확인 후 다시 등록하세요.";
                                PSH_Globals.SBO_Application.MessageBox(errMessage);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY035A_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (!string.IsNullOrEmpty(oForm.Items.Item("BeforKm").Specific.Value.ToString().Trim()) || !string.IsNullOrEmpty(oForm.Items.Item("AfterKm").Specific.Value.ToString().Trim()))
                            {
                                sQry = " UPDATE [@PH_PY035A] SET U_RegCls ='03' WHERE DocEntry ='" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }

                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {

                        }
                    }
                    if (pVal.ItemUID == "BtnPrint")
                    {
                        if (PH_PY035A_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        report_print_035(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
                    }
                }

                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY035_FormItemEnabled();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PH_PY035_MTX01(1);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// FormMenuEvent
        /// <summary>
        /// 메뉴이벤트
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
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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
                            PH_PY035_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PH_PY035_FormItemEnabled();
                            PH_PY035_SetDocEntry();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PH_PY035_FormItemEnabled();
                            break;
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}

