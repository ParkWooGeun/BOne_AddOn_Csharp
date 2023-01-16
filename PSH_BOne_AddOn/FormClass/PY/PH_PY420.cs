using System;
using SAPbouiCOM;
using System.Text;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.ExportFile;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연말정산 PDF 업로드
    /// </summary>
    internal class PH_PY420 : PSH_BaseClass
    {
        public string oFormUniqueID;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01">문서번호</param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY420.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue.ToString() + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int loopCount = 1; loopCount <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); loopCount++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[loopCount - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[loopCount - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY420_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY420");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(ref strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY420_Create_Items();
                PH_PY420_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY420_Create_Items()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                //성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //파일경로
                oForm.DataSources.UserDataSources.Add("Comments", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 200);
                oForm.Items.Item("Comments").Specific.DataBind.SetBound(true, "", "Comments");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY420_Create_Items_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY420_FormItemEnabled()
        {
            try
            {
                oForm.EnableMenu("1282", true);  // 문서추가
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("MSTCOD").Specific.Value = "";
                    oForm.Items.Item("FullName").Specific.Value = "";
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY420_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PDF자료를 공유폴더에 Xml 파일로 저장
        /// </summary>
        [STAThread]
        private void PH_PY420_PDF_Upload()
        {
            int fileSize = 0;
            string XmlText;
            string oFilePath;
            string MSTCOD;
            string password;
            string UserID;
            MSXML2.DOMDocument xmldoc = new MSXML2.DOMDocument();
            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            try
            {
                commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("PDF Files", "*.PDF"));

                if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    oForm.Items.Item("Comments").Specific.Value = commonOpenFileDialog.FileName;
                }

                string filePath = oForm.Items.Item("Comments").Specific.Value;
                password = oForm.Items.Item("Password").Specific.Value;
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                string strXML = "XML";

                if (string.IsNullOrEmpty(filePath))
                {
                    return;
                }

                fileSize = EXPFile.NTS_GetFileSize(filePath, password, strXML, 0);

                if (fileSize > 0)
                {
                    byte[] buf = new byte[fileSize];
                    fileSize = EXPFile.NTS_GetFileBuf(filePath, password, strXML, buf, 0);
                    if (fileSize > 0)
                    {
                        string strBuf = Encoding.UTF8.GetString(buf);
                        strBuf.Replace("\n", "\r\n");
                        XmlText = strBuf;
                        xmldoc.loadXML(XmlText);
                        oFilePath = @"\\191.1.1.220\pdf\";

                        UserID = MSTCOD + "_PDFtoXML.xml";
                        xmldoc.save(oFilePath + UserID);
                        PSH_Globals.SBO_Application.StatusBar.SetText("PDF파일이 변환 되었습니다. 아래 계산버튼을 클릭하세요. ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }
            }
            catch (Exception ex)
            {
                if (fileSize == -10)
                {
                    PSH_Globals.SBO_Application.MessageBox("파일이 없거나 손상된 PDF 파일입니다.");
                }
                else if (fileSize == -11)
                {
                    PSH_Globals.SBO_Application.MessageBox("국세청에서 발급된 전자문서가 아닙니다.");
                }
                else if (fileSize == -13)
                {
                    PSH_Globals.SBO_Application.MessageBox("추출용 버퍼가 유효하지 않습니다.");
                }
                else if (fileSize == -200)
                {
                    PSH_Globals.SBO_Application.MessageBox("비밀번호가 틀립니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY420_PDF_Upload : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY420_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                {
                    errMessage = "사업장은 필수입니다.";
                    throw new Exception();
                }
               
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.Trim()))
                {
                    errMessage = "년도는 필수입니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.Trim()))
                {
                    errMessage = "사번은 필수입니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("FullName").Specific.Value.Trim()))
                {
                    errMessage = "이름은 필수입니다.";
                    throw new Exception();
                }

                sQry = "select count(*)  from [p_seoybase] where div in (10) and yyyy = '" + oForm.Items.Item("Year").Specific.Value.Trim() + "' and sabun ='" + oForm.Items.Item("MSTCOD").Specific.Value.Trim() + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    errMessage = "필수 : 기본공제 등록";
                    errMessage += Environment.NewLine;
                    errMessage += "선택 : 추가공제, 기본공제제외자";
                    errMessage += Environment.NewLine;
                    errMessage += "먼저 등록후 PDF 업로드 실행하세요.";

                    throw new Exception();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// Xml 자료를 Table에 Insert 
        /// </summary>
        private bool PH_PY420_PDFtoTable_insert()
        {
            bool returnValue = false;
            string sQry;
            string oFilePath;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oFilePath = "\\\\" + PSH_Globals.SP_ODBC_IP + "\\pdf\\";

                sQry = "INSERT INTO TBL_XML(CREATEDATE, BPLID, yyyy, MSTCOD, XMLDATA)";
                sQry += " SELECT GETDATE() AS CREATEDATE,'" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "','" + oForm.Items.Item("Year").Specific.Value;
                sQry += "','" + oForm.Items.Item("MSTCOD").Specific.Value + "' AS MSTCOD, * FROM OPENROWSET (";
                sQry += "BULK '" + oFilePath + oForm.Items.Item("MSTCOD").Specific.Value + "_PDFtoXML.xml', SINGLE_BLOB) AS x";
                oRecordSet.DoQuery(sQry);

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY420_PDFtoTable_insert : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return returnValue;
        }

        /// <summary>
        /// XML 실행
        /// </summary>
        private void PH_PY420_Covert_execution()
        {
            string sQry;
            string CLTCOD;
            string Year;
            string MSTCOD;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                sQry = "Exec PH_PY420_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                sQry = "Exec PH_PY420_02 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("PDF자료가 등록 되었습니다. ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY420_Covert_execution : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
        /// VALIDATE 이벤트
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                sQry = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " and Code = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";

                                oRecordSet.DoQuery(sQry);
                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                break;

                            case "FullName":
                                sQry = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += " And U_status <> '5'"; // 퇴사자 제외
                                sQry += " and U_FullName = '" + oForm.Items.Item("FullName").Specific.Value + "'";

                                oRecordSet.DoQuery(sQry);
                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
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
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        if (PH_PY420_DataValidCheck() == true)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY420_PDF_Upload);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PH_PY420_DataValidCheck() == true)
                        {
                            if (PH_PY420_PDFtoTable_insert() == true)
                            {
                                PH_PY420_Covert_execution();
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
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
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            break;
                        case "1282": //문서추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "1293": // 행삭제
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
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}