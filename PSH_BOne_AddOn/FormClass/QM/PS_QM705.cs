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
using System.Text.RegularExpressions;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전송
    /// </summary>
    internal class PS_QM705 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_QM705H; //등록헤더
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_QM705M;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM705.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_QM705_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_QM705");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                oForm.Freeze(true);
                PS_QM705_CreateItems();
                PS_QM705_ComboBox_Setting();
                PS_QM705_Initialization();
                PS_QM705_AddMatrixRow(0, true);
                PS_QM705_AddMatrixRowM(0, true);
            }
            catch (Exception ex)
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
        private void PS_QM705_CreateItems()
        {
            try
            {
                oDS_PS_QM705H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_QM705M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("oMat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 메트릭스 개체 할당
                oMat02 = oForm.Items.Item("oMat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();


                oForm.DataSources.UserDataSources.Add("DocDatefr", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDatefr").Specific.DataBind.SetBound(true, "", "DocDatefr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("DocDateto", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateto").Specific.DataBind.SetBound(true, "", "DocDateto");
                oForm.DataSources.UserDataSources.Item("DocDateto").Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_QM705_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("CLTCOD").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_QM705_Initialization
        /// </summary>
        private void PS_QM705_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PS_QM705_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_QM705H.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_QM705H.Offset = oRow;
                oDS_PS_QM705H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM705H_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PS_QM705_AddMatrixRowM(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_QM705M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_QM705M.Offset = oRow;
                oDS_PS_QM705M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM705M_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_QM705_LoadData()
        {
            string sQry;
            string CLTCOD;
            string MSTCOD;
            string DocDateFr;
            string DocDateTo;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PS_QM705H.Clear();

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                DocDateFr = oForm.Items.Item("DocDatefr").Specific.Value.ToString().Trim();
                DocDateTo = oForm.Items.Item("DocDateto").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //검사자임

                sQry = "EXEC PS_QM702_02 '" + CLTCOD + "','" + DocDateFr + "','" + DocDateTo + "','" + MSTCOD + "'";
                oRecordSet01.DoQuery(sQry);

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_QM705H.Size)
                    {
                        oDS_PS_QM705H.InsertRecord((i));
                    }
                    oMat01.AddRow();
                    oDS_PS_QM705H.Offset = i;
                    oDS_PS_QM705H.SetValue("U_LineNum", i, Convert.ToString(i + 1));  // 순번
                    oDS_PS_QM705H.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocDate").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("Gubun").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("WorkNum").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("WorkDate").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("WorkCode").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("WorkName").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("BZZadQty").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("BadCode").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("BadeNote").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("verdict").Value.ToString().Trim());
                    oDS_PS_QM705H.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY705_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PDF만들기
        /// </summary>
        [STAThread]
        private bool Make_PDF_File(String p_DocEntry, string p_Gobun)
        {
            bool ReturnValue = false;
            string WinTitle;
            string ReportName = String.Empty;
            string Main_Folder;
            string ExportString;
            string Incom_Pic_Path;
            string filename;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (p_Gobun == "외주")
                {
                    filename = "_Out.bmp";
                }
                else
                {
                    filename = "_In.bmp";
                }
                Incom_Pic_Path = @"\\191.1.1.220\Incom_Pic\";

                if (File.Exists(Incom_Pic_Path + p_DocEntry + filename))
                {
                    if (File.Exists(Incom_Pic_Path + "PIC.bmp") == true)
                    {
                        File.Delete(Incom_Pic_Path + "PIC.bmp");
                        File.Copy(Incom_Pic_Path + p_DocEntry + filename, Incom_Pic_Path + "PIC.bmp");
                    }
                    else
                    {
                        File.Copy(Incom_Pic_Path + p_DocEntry + filename, Incom_Pic_Path + "PIC.bmp");
                    }
                }
                else
                {
                    File.Delete(Incom_Pic_Path + "PIC.bmp");
                    File.Copy(Incom_Pic_Path + "NULL.bmp", Incom_Pic_Path + "PIC.bmp");
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackSub1ReportParameter = new List<PSH_DataPackClass>(); //서브레포트 그대로날리는변수 

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", p_DocEntry));

                Main_Folder = @"C:\PSH_부적합전송";
                Dir_Exists(Main_Folder);
                ExportString = Main_Folder + @"\" + "풍산홀딩스부적합보고서" + p_Gobun + p_DocEntry + ".pdf";

                if (p_Gobun == "외주")
                {
                    WinTitle = "[PS_QM705] 외주 부적합 자재 통보서";
                    ReportName = "PS_QM702_01.rpt";
                    dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@DocEntry", p_DocEntry, "PS_QM702_04"));

                }
                else
                {
                    WinTitle = "[PS_QM705] 자체 부적합 자재 보고서";
                    ReportName = "PS_QM702_02.rpt";
                    dataPackSub1ReportParameter.Add(new PSH_DataPackClass("@DocEntry", p_DocEntry, "SUB702_06"));
                }

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackSub1ReportParameter, ExportString);

                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(ExportString, PdfDocumentOpenMode.Modify);

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                //securitySettings.UserPassword = "manager";   //개개인암호
                //securitySettings.OwnerPassword = psgovID;    //마스터암호

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
        /// <param name="p_DocEntry"></param>
        /// <param name="p_Address"></param>
        /// <returns></returns>
        private bool Send_EMail(string p_DocEntry, string p_Address, string p_Gobun)
        {
            bool ReturnValue = false;
            string Main_Folder;
            string sQry;
            string errMessage = string.Empty;
            string signature = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (p_Gobun == "외주")
                {
                    sQry = " SELECT U_Comments FROM [@PS_QM701H] WHERE DocEntry ='" + p_DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);
                }
                else
                {
                    sQry = " SELECT U_Comments FROM [@PS_QM703H] WHERE DocEntry ='" + p_DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);
                }
                Main_Folder = @"C:\PSH_부적합전송";
                MsOutlook.Application outlookApp = new MsOutlook.Application();
                if (outlookApp == null)
                {
                    throw new Exception();
                }



                //string signatureFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Signatures\test.htm");
                //if (File.Exists(signatureFilePath))
                //{
                //    // 파일에서 서명 읽어오기
                //    string signatureContent = File.ReadAllText(signatureFilePath, System.Text.Encoding.UTF8);
                //    // HTML 태그 제거
                //    string strippedSignature = Regex.Replace(signatureContent, "<.*?>", string.Empty);
                //    signature = strippedSignature;
                //}
                //else
                //{
                //    errMessage = "서명 파일이 존재하지 않습니다.";
                //    throw new Exception();
                //}

                string signatureFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Signatures");
                DirectoryInfo diInfo = new DirectoryInfo(signatureFilePath);
                if (diInfo.Exists)
                {
                    FileInfo[] fiSignature = diInfo.GetFiles("*.txt");

                    if (fiSignature.Length > 0)
                    {
                        StreamReader sr = new StreamReader(fiSignature[0].FullName, System.Text.Encoding.UTF8);
                        signature = sr.ReadToEnd();

                        if (!string.IsNullOrEmpty(signature))
                        {
                            string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                            signature = signature.Replace(fileName + "_files/", signatureFilePath + "/" + fileName + "_files/");
                        }
                    }
                }
                MsOutlook.MailItem mail = (MsOutlook.MailItem)outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);
                mail.Subject = oForm.Items.Item("Subject").Specific.Value.ToString().Trim();
                mail.HTMLBody = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body>" + oForm.Items.Item("SBody").Specific.Value.ToString().Trim() + Environment.NewLine + signature + "</ body ></ html >";
                mail.To = p_Address;
                MsOutlook.Attachment oAttach1 = mail.Attachments.Add(Main_Folder + @"\" + "풍산홀딩스부적합보고서" + p_Gobun + p_DocEntry + ".pdf");
                if (!string.IsNullOrEmpty(oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim()))
                {
                    MsOutlook.Attachment oAttach2 = mail.Attachments.Add(oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                }
                mail.Send();

                mail = null;
                outlookApp = null;
                ReturnValue = true;
            }
            catch (System.Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Send_EMail_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
        /// MATRIX_LOAD 이벤트
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
                    PS_QM705_AddMatrixRowM(oMat02.VisualRowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            string sQry1;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == true)
                {
                   if (pVal.ItemUID == "btn_search")
                    {
                        PS_QM705_LoadData();
                    }
                    else if (pVal.ItemUID == "btn_send")
                    {
                        string SDocEntry = oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim();
                        string SGoBun = oForm.Items.Item("SGoBun").Specific.Value.ToString().Trim();
                        if (string.IsNullOrEmpty(oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim()))
                        {
                            errMessage = "결재완료된 문서번호 선택은 필수입니다.";
                            throw new Exception();
                        }
                        if (string.IsNullOrEmpty(oForm.Items.Item("Subject").Specific.Value.ToString().Trim()))
                        {
                            errMessage = "제목을 입력해주세요.";
                            throw new Exception();
                        }
                        if (string.IsNullOrEmpty(oForm.Items.Item("SBody").Specific.Value.ToString().Trim()))
                        {
                            errMessage = "본문을 입력해주세요.";
                            throw new Exception();
                        }
                       
                        if ( Make_PDF_File(SDocEntry, SGoBun) == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if(oMat02.VisualRowCount-1 == 0)
                        {
                            errMessage = "메일을 보낼 주소가 없습니다. 확인해주세요.";
                            throw new Exception();
                        }
                        else
                        {
                            for (int i = 0; i <= oMat02.VisualRowCount - 1; i++)
                            {
                                int k = 0;
                                sQry = "SELECT U_eMail FROM [@PS_QM700L] WHERE U_UseYN = 'Y'AND Code ='" + oDS_PS_QM705M.GetValue("U_ColReg01", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                for (int j = 0; j <= oRecordSet01.RecordCount - 1; j++)
                                {
                                    if (Send_EMail(SDocEntry, oRecordSet01.Fields.Item("U_eMail").Value.ToString().Trim(), SGoBun) == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                    oDS_PS_QM705M.SetValue("U_ColReg03", i, "Y");
                                    oRecordSet01.MoveNext();
                                    k++;
                                }
                                if( k != 0)
                                {
                                    sQry1 = "Insert into Z_PS_QM702 values ('" + SGoBun + "','" + SDocEntry + "','";
                                    sQry1 += oDS_PS_QM705M.GetValue("U_ColReg01", i).ToString().Trim() + "','" + oDS_PS_QM705M.GetValue("U_ColReg02", i).ToString().Trim() + "',GETDATE())";
                                    oRecordSet02.DoQuery(sQry1);
                                }
                            }
                        }
                        oMat02.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        errMessage = "전송이 완료되었습니다.";
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                   else if (pVal.ItemUID == "oMat02")
                    {
                        if (pVal.ColUID == "Code")
                        {
                            if (string.IsNullOrEmpty(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
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
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "oMat02")
                        {
                            if (pVal.ColUID == "Code")
                            {
                                oMat02.FlushToDataSource();
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_QM705M.GetValue("U_ColReg01" , pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_QM705_AddMatrixRowM(pVal.Row, false);
                                }
                                oMat02.LoadFromDataSource();
                            }
                        }
                       
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "MSTCOD")
                        {
                            oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value + "'", ""); //검사자
                        }
                        if (pVal.ItemUID == "oMat02")
                        {
                            if (pVal.ColUID == "Code")
                            {
                                sQry = "select Name from[@PS_QM700H] where Code = '" + oMat02.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value + "'";
                                RecordSet01.DoQuery(sQry);
                                oMat02.Columns.Item("Name").Cells.Item(pVal.Row).Specific.Value = RecordSet01.Fields.Item(0).Value;                                      
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM705H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM705M);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "oMat01")
                    {
                        sQry = "안녕하십니까 풍산홀딩스 품질보증팀 이지윤 사원입니다.<br>\n";
                        sQry += "   귀사의 일익 번창하심을 기원합니다.<br>\n";
                        sQry += "   첨부와 같이 부적합자재통보서를 발송합니다.<br>\n";
                        sQry += "   통보서에 명시한 기한내에 현품 처리 및 불량 발생원인 대책 회신 바랍니다.<br>\n";
                        oForm.Items.Item("SDocEntry").Specific.Value = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
                        oForm.Items.Item("SGoBun").Specific.Value = oMat01.Columns.Item("Gubun").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
                        oForm.Items.Item("Subject").Specific.Value = "(풍산홀딩스) 부적합자재통보서 송부의 건";
                        oForm.Items.Item("SBody").Specific.Value = sQry;

                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        oDS_PS_QM705M.Clear();
                        PS_QM705_AddMatrixRowM(0, false);
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
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                        case "1282": //추가
                            PS_QM705_AddMatrixRowM(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
