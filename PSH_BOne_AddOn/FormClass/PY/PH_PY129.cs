using System;
using System.IO;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 개인별 퇴직연금(DC형) 계산
    /// </summary>
    internal class PH_PY129 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY129A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY129B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY129.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY129_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY129");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY129_CreateItems();
                PH_PY129_EnableMenus();
                PH_PY129_SetDocument(oFormDocEntry);
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
        private void PH_PY129_CreateItems()
        {
            string CLTCOD;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY129A = oForm.DataSources.DBDataSources.Item("@PH_PY129A");
                oDS_PH_PY129B = oForm.DataSources.DBDataSources.Item("@PH_PY129B");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();
                
                CLTCOD = dataHelpClass.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'","");
                oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY129_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY129_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY129_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY129_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Btn01").Visible = true;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("Btn01").Visible = true;
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Btn01").Visible = false;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY129_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY129_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY129_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY129_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY129_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY129_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY129B.GetValue("U_Date", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY129B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY129B.InsertRecord((oRow));
                        }
                        oDS_PH_PY129B.Offset = oRow;
                        oDS_PH_PY129B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY129B.SetValue("U_Date", oRow, "");
                        oDS_PH_PY129B.SetValue("U_Type", oRow, "");
                        oDS_PH_PY129B.SetValue("U_Comments", oRow, "0");
                        oDS_PH_PY129B.SetValue("U_Close", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY129B.Offset = oRow - 1;
                        oDS_PH_PY129B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY129B.SetValue("U_Date", oRow - 1, "");
                        oDS_PH_PY129B.SetValue("U_Type", oRow - 1, "");
                        oDS_PH_PY129B.SetValue("U_Comments", oRow - 1, "0");
                        oDS_PH_PY129B.SetValue("U_Close", oRow - 1, "0");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY129B.Offset = oRow;
                    oDS_PH_PY129B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY129B.SetValue("U_Date", oRow, "");
                    oDS_PH_PY129B.SetValue("U_Type", oRow, "");
                    oDS_PH_PY129B.SetValue("U_Comments", oRow, "0");
                    oDS_PH_PY129B.SetValue("U_Close", oRow, "0");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY129_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY129_DataValidCheck(string ChkYN)
        {
            bool returnValue = false;
            string sQry;
            string tCode;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (ChkYN == "Y")
                {
                    if (string.IsNullOrEmpty(oDS_PH_PY129A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return returnValue;
                    }
                }

                if (string.IsNullOrEmpty(oDS_PH_PY129A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("기준년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //코드,이름 저장
                tCode = oDS_PH_PY129A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY129A.GetValue("U_CLTCOD", 0).ToString().Trim();
                oDS_PH_PY129A.SetValue("Code", 0, tCode);
                oDS_PH_PY129A.SetValue("Name", 0, tCode);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //데이터 중복 체크
                    sQry = "SELECT Code FROM [@PH_PY129A] WHERE Code = '" + tCode + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 데이터가 존재합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return returnValue;
                    }
                }

                if (ChkYN == "Y")
                {
                    if (oMat1.VisualRowCount == 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 없습니다. 확인바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return returnValue;
                    }
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY129_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }

            return returnValue;
        }

        /// <summary>
        /// DataFind : 자료 조회
        /// </summary>
        private void PH_PY129_LoadData()
        {
            int i;
            string sQry;
            string CLTCOD;
            string YM;
            string MSTCOD;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.Value;
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                sQry = "EXEC PH_PY129_01 '" + CLTCOD + "', '" + YM + "', '" + MSTCOD + "', '10', ''";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY129B.InsertRecord(i);
                    }
                    oDS_PH_PY129B.Offset = i;
                    oDS_PH_PY129B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY129B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_FullName", i, oRecordSet.Fields.Item("FullName").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_BirthDat", i, oRecordSet.Fields.Item("BirthDat").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_TeamName", i, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_RspCode", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_RspName", i, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_ClsCode", i, oRecordSet.Fields.Item("ClsCode").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_ClsName", i, oRecordSet.Fields.Item("ClsName").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_InpDat", i, oRecordSet.Fields.Item("InpDat").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_GrpDat", i, oRecordSet.Fields.Item("GrpDat").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_RETDAT", i, oRecordSet.Fields.Item("RETDAT").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_JIGTYP", i, oRecordSet.Fields.Item("JIGTYP").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_JIGTYPNM", i, oRecordSet.Fields.Item("JIGTYPNM").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_JIGCOD", i, oRecordSet.Fields.Item("JIGCOD").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_JIGCODNM", i, oRecordSet.Fields.Item("JIGCODNM").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_PAYTYP", i, oRecordSet.Fields.Item("PAYTYP").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_PAYTYPNM", i, oRecordSet.Fields.Item("PAYTYPNM").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_TOTPAY", i, oRecordSet.Fields.Item("TOTPAY").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_MMCnt", i, oRecordSet.Fields.Item("MMCnt").Value.ToString().Trim());
                    oDS_PH_PY129B.SetValue("U_AVGPAY", i, oRecordSet.Fields.Item("AVGPAY").Value.ToString().Trim());

                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY129_LoadData_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY129_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string CLTCOD;
            string Code;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                Code = oForm.Items.Item("Code").Specific.Value.Trim();

                WinTitle = "[PH_PY129_05] D/C형 퇴직연금계산내역서";
                ReportName = "PH_PY129_05.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@CODE", Code)); //지급년월

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY129_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PDF만들기
        /// </summary>
        [STAThread]
        private bool Make_PDF_File(String p_MSTCOD)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName = String.Empty;
            string CLTCOD;
            string Code;
            string Main_Folder;
            string Sub_Folder1;
            string sQry1;
            string sQry;
            string ExportString;
            string psgovID;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                Code = oForm.Items.Item("Code").Specific.Value.Trim();

                WinTitle = "[PH_PY129] D/C형 퇴직연금계산내역서";
                ReportName = "PH_PY129_06.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@CODE", Code)); //지급년월
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", p_MSTCOD)); //지급년월


                Main_Folder = @"C:\PSH_DC형 퇴직연금계산내역서";
                Sub_Folder1 = @"C:\PSH_DC형 퇴직연금계산내역서\" + Code + "";
                //Sub_Folder2 = @"C:\PSH_D/C형 퇴직연금계산내역서\" + Code + @"\" + CLTCOD + "";

                Dir_Exists(Main_Folder);
                Dir_Exists(Sub_Folder1);
                //Dir_Exists(Sub_Folder2);

                sQry1 = " exec [PH_PY129_06] '" + CLTCOD + "','" + Code + "','" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry1);

                ExportString = Sub_Folder1 + @"\" + p_MSTCOD + ".pdf";

                sQry = "Select RIGHT(U_govID,7) From [@PH_PY001A]";
                sQry += "WHERE  Code ='" + p_MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);
                psgovID = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, ExportString, 100);

                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(ExportString, PdfDocumentOpenMode.Modify);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                securitySettings.UserPassword = "manager";   //개개인암호
                securitySettings.OwnerPassword = psgovID;    //마스터암호

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // PDF문서 저장
                document.Save(ExportString);

                sQry = "Update [@PH_PY129B] Set U_SaveYN = 'Y' Where U_MSTCOD = '" + p_MSTCOD + "' And Code = '" + Code + "'";
                oRecordSet01.DoQuery(sQry);

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// 디렉토리 체크, 폴더 생성
        /// </summary>
        /// <param name="strDirName">경로</param>
        /// <returns></returns>
        private int Dir_Exists(string strDirName)
        {
            int ReturnValue = 0;

            try
            {
                DirectoryInfo di = new DirectoryInfo(strDirName); //DirectoryInfo 생성
                //DirectoryInfo.Exists로 폴더 존재유무 확인
                if (di.Exists)
                {
                    ReturnValue = 1;
                }
                else
                {
                    di.Create();
                    ReturnValue = 0;
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Make_PDF_File_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return ReturnValue;
        }

        /// <summary>
        /// Send_EMail
        /// </summary>
        /// <param name="p_MSTCOD"></param>
        /// <param name="p_Version"></param>
        /// <returns></returns>
        private bool Send_EMail(string p_MSTCOD)
        {
            bool ReturnValue = false;
            string strToAddress;
            string strSubject;
            string strBody;
            string Sub_Folder1;
            string sQry;
            string Code;
            string MSTCOD;
            string CLTCOD;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MSTCOD = p_MSTCOD;
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                Code = oForm.Items.Item("Code").Specific.Value.Trim();

                Sub_Folder1 = @"C:\PSH_DC형 퇴직연금계산내역서\" + Code + "";

                sQry = "Select U_Subject, U_Body From [@PH_PY129A] Where Code = '" + Code + "'";
                oRecordSet01.DoQuery(sQry);
                strSubject = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                strBody = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                sQry = "SELECT b.U_eMail FROM [@PH_PY129B] a inner join [@PH_PY001A] B ON A.U_MSTCOD = B.Code WHERE a.U_MSTCOD = '" + MSTCOD + "' AND a.Code = '" + Code + "'";
                oRecordSet01.DoQuery(sQry);
                strToAddress = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                //mail.From = new MailAddress("dakkorea1@gmail.com");
                MsOutlook.Application outlookApp = new MsOutlook.Application();
                if (outlookApp == null)
                {
                    throw new Exception();
                }
                MsOutlook.MailItem mail = (MsOutlook.MailItem)outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);

                mail.Subject = strSubject;
                mail.HTMLBody = strBody;
                mail.To = strToAddress;
                MsOutlook.Attachment oAttach = mail.Attachments.Add(Sub_Folder1 + @"\" + p_MSTCOD + ".pdf");
                mail.Send();

                mail = null;
                outlookApp = null;

                sQry = "Update [@PH_PY129B] Set U_SendYN = 'Y' Where U_MSTCOD = '" + p_MSTCOD + "' And Code = '" + Code + "'";
                oRecordSet01.DoQuery(sQry);

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(Sub_Folder3 + @"\" + p_MSTCOD + "_개인별급여명세서_" + STDYER + "" + STDMON + ".pdf");

                //원래코드시작
                //SmtpClient smtp = new SmtpClient("smtp.naver.com");
                //SmtpClient smtp = new SmtpClient("pscsn.poongsan.co.kr");
                //SmtpClient smtp = new SmtpClient("smtp.office365.com");
                //SmtpClient smtp = new SmtpClient("smtp.gmail.com");

                //smtp.Port = 587; //네이버
                //smtp.Port = 25; //풍산
                //smtp.UseDefaultCredentials = true;
                //smtp.EnableSsl = true;
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //smtp.Timeout = 20000;

                //smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;  //Naver 인 경우
                //smtp.Credentials = new NetworkCredential("2220501", "p2220501!"); //address, PW
                //smtp.Credentials = new NetworkCredential("wgpark@poongsan.co.kr", "1q2w3e4r)*"); //address, PW
                //smtp.Credentials = new NetworkCredential("dakkorea1@gmail.com", "dak440310*"); //address, PW

                //smtp.Send(mail);
                //원래코드 끝

                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Send_EMail_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return ReturnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string CLTCOD;
            string YM;
            string Code;
            string MSTCOD;
            string p_MSTCOD;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY129_DataValidCheck("Y") == false)
                            {
                                BubbleEvent = false;
                            }

                            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                            Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                            sQry = "EXEC PH_PY129_01 '" + CLTCOD + "', '" + YM + "', '" + MSTCOD + "', '20','" + Code + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PH_PY129_DataValidCheck("N") == false)
                        {
                        }
                        else
                        {
                            PH_PY129_LoadData();
                        }
                    }
                    if (pVal.ItemUID == "btn_save")
                    {
                        if (PH_PY129_DataValidCheck("N") == false)
                        {
                        }
                        else
                        {
                           //CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value;
                            Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
                            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("PDF 파일 생성 시작!", 50, false);
                            for (int i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                if (!string.IsNullOrEmpty(oDS_PH_PY129B.GetValue("U_MSTCOD", i).ToString().Trim()))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY129B.GetValue("U_SendYN", i).ToString().Trim()))
                                    {
                                        p_MSTCOD = oDS_PH_PY129B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        if (Make_PDF_File(p_MSTCOD) == false)
                                        {
                                            errMessage = "PDF저장이 완료되지 않았습니다.";
                                            throw new Exception();
                                        }
                                    }
                                }
                                ProgressBar01.Value += 1;
                                ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat1.VisualRowCount) + "건 PDF 파일 생성 중...!";
                            }
                            ProgressBar01.Stop();

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PH_PY129_FormItemEnabled();
                            oForm.Items.Item("Code").Specific.Value = Code;
                            //oForm.Items.Item("CLTCOD").Specific.Selected.Value = CLTCOD;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                    if (pVal.ItemUID == "btn_send")
                    {
                        if (PH_PY129_DataValidCheck("N") == false)
                        {
                        }
                        else
                        {
                            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
                            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("eMail 메일전송", 50, false);
                            oMat1.FlushToDataSource();
                            for (int i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                 if (!string.IsNullOrEmpty(oDS_PH_PY129B.GetValue("U_SaveYN", i).ToString().Trim()))
                                    {
                                        p_MSTCOD = oDS_PH_PY129B.GetValue("U_MSTCOD", i).ToString().Trim();
                                        if (Send_EMail(p_MSTCOD) == false)//사번
                                        {
                                            errMessage = "전송 중 오류가 발생했습니다.";
                                            throw new Exception();
                                        }
                                    }
                                ProgressBar01.Value += 1;
                                ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat1.VisualRowCount) + "건 eMail전송중...!";
                            }
                            ProgressBar01.Stop();

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PH_PY129_FormItemEnabled();
                            oForm.Items.Item("YM").Specific.Value = Code.Substring(0, 6);
                            oForm.Items.Item("CLTCOD").Specific.Select.Value = CLTCOD;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                    if (pVal.ItemUID == "Btn_prt")
                    {
                        Code = oForm.Items.Item("Code").Specific.Value;
                        if (!string.IsNullOrEmpty(Code))
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY129_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("조회후 출력 하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY129_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY129_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
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
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
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
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD": 
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" +oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'"; 
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("MSTNAM").Specific.String = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim(); //사원명
                                break;
                        }
                    }
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
                    oMat1.LoadFromDataSource();
                    PH_PY129_FormItemEnabled();
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY129A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY129B);
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
            string Code;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            Code = oForm.Items.Item("Code").Specific.Value;
                            sQry = "Delete From Z_PH_PY129 Where Code = '" + Code + "'";
                            oRecordSet.DoQuery(sQry);
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY129A", "Code"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY129_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY129_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY129_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY129_FormItemEnabled();
                            break;
                        case "1293": //행삭제
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
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
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