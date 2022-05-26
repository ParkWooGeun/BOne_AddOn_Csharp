//using System;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 프로그램 개발요청 관리
//    /// </summary>
//    internal class PS_AD103 : PSH_BaseClass
//    {
//        private string oFormUniqueID;
//        private SAPbouiCOM.Matrix oMat01;
//        private SAPbouiCOM.DBDataSource oDS_PS_AD103H; //등록헤더
//        private SAPbouiCOM.DBDataSource oDS_PS_AD103L; //등록라인

//        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
//        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

//        private string oDocEntry01;

//        private SAPbouiCOM.BoFormMode oFormMode01;

//        /// <summary>
//        /// Form 호출
//        /// </summary>
//        /// <param name="oFromDocEntry01"></param>
//        public override void LoadForm(string oFromDocEntry01)
//        {
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_AD103.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PS_AD103_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PS_AD103");

//                string strXml = null;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                oForm.DataBrowser.BrowseBy = "DocEntry";

//                oForm.Freeze(true);
//                PS_AD103_CreateItems();
//                PS_AD103_ComboBox_Setting();
//                PS_AD103_CF_ChooseFromList();
//                PS_AD103_EnableMenus();
//                PS_AD103_SetDocument(oFromDocEntry01);
//                PS_AD103_FormResize();

//                oForm.EnableMenu(("1283"), false); //삭제
//                oForm.EnableMenu(("1287"), false); //복제
//                oForm.EnableMenu(("1286"), false); //닫기
//                oForm.EnableMenu(("1284"), true); //취소
//                oForm.EnableMenu(("1293"), true); //행삭제
//                oForm.EnableMenu(("1299"), false); //행닫기
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
//            }
//        }

//        /// <summary>
//        /// 화면 Item 생성
//        /// </summary>
//        private void PS_AD103_CreateItems()
//        {
//            try
//            {
//                oDS_PS_AD103H = oForm.DataSources.DBDataSources.Item("@PS_AD103H");
//                oDS_PS_AD103L = oForm.DataSources.DBDataSources.Item("@PS_AD103L");
//                oMat01 = oForm.Items.Item("Mat01").Specific;

//                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// Combobox 설정
//        /// </summary>
//        private void PS_AD103_ComboBox_Setting()
//        {
//            string sQry;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

//                //요청구분(Matrix)
//                sQry = "        SELECT      U_Minor, ";
//                sQry += "             U_CdName ";
//                sQry += " FROM        [@PS_SY001L] ";
//                sQry += " WHERE       Code = 'A002'";
//                sQry += "             AND ISNULL(U_UseYN, 'Y') = 'Y'";
//                sQry += " ORDER BY    U_Seq";

//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ReqType"), sQry,"","");

//                //Action(Matrix)
//                sQry = "        SELECT      U_Minor, ";
//                sQry += "             U_CdName ";
//                sQry += " FROM        [@PS_SY001L] ";
//                sQry += " WHERE       Code = 'A009'";
//                sQry += "             AND ISNULL(U_UseYN, 'Y') = 'Y'";
//                sQry += " ORDER BY    U_Seq";

//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Action"), sQry, "", "");
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// 처리가능한 Action인지 검사
//        /// </summary>
//        /// <param name="ValidateType"></param>
//        /// <returns></returns>
//        private bool PS_AD103_Validate(string ValidateType)
//        {
//            bool returnValue = false;
//            int i;
//            int j;
//            string query01;
//            bool Exist;
//            string errCode = string.Empty;
//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (ValidateType == "수정")
//                {
//                }
//                else if (ValidateType == "RowDelete")
//                {
//                    ////행삭제전 행삭제가능여부검사
//                    ////추가,수정모드일때행삭제가능검사
//                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                    {
//                        ////새로추가된 행인경우, 삭제하여도 무방하다
//                        if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.VALUE)))
//                        {
//                        }
//                        else
//                        {
//                            if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_AD104L] WHERE U_BasEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "-" + oMat01.Columns.Item("LineID").Cells.Item(oLastColRow01).Specific.VALUE + "'", 0, 1)) > 0)
//                            {
//                                MDC_Com.MDC_GF_Message(ref "개발품의서가 등록된 행입니다. 삭제할수 없습니다.", ref "W");
//                                functionReturnValue = false;
//                                goto PS_AD103_Validate_Exit;
//                            }
//                        }
//                    }
//                }
//                returnValue = true;
//            }
//            catch (Exception ex)
//            {
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }
//            return returnValue;
//        }

//        /// <summary>
//        /// EnableMenus
//        /// </summary>
//        private void PS_AD103_EnableMenus()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false ,false);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// SetDocument
//        /// </summary>
//        /// <param name="oFromDocEntry01">DocEntry</param>
//        private void PS_AD103_SetDocument(string oFromDocEntry01)
//        {
//            try
//            {
//                if ((string.IsNullOrEmpty(oFromDocEntry01)))
//                {
//                    PS_AD103_FormItemEnabled();
//                    PS_AD103_AddMatrixRow(0, true);
//                }
//                else
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                    PS_AD103_FormItemEnabled();
//                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// 모드에 따른 아이템 설정
//        /// </summary>
//        private void PS_AD103_FormItemEnabled()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                    oForm.Items.Item("DocEntry").Enabled = false;
//                    oForm.Items.Item("Mat01").Enabled = true;

//                    PS_AD103_FormClear();
//                    oForm.EnableMenu("1281", true); //찾기
//                    oForm.EnableMenu("1282", false); //추가

//                    //사용자별 사업장 세팅
//                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//                {
//                    oForm.Items.Item("DocEntry").Specific.VALUE = "";
//                    oForm.Items.Item("DocEntry").Enabled = true;
//                    oForm.Items.Item("Mat01").Enabled = false;

//                    oForm.EnableMenu("1281", false); ////찾기
//                    oForm.EnableMenu("1282", true); ////추가

//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//                {
//                    oForm.Items.Item("DocEntry").Enabled = false;
//                    oForm.Items.Item("Mat01").Enabled = true;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// PS_AD103_AddMatrixRow
//        /// </summary>
//        /// <param name="oRow">행 번호</param>
//        /// <param name="RowIserted">행 추가 여부</param>
//        private void PS_AD103_AddMatrixRow(int oRow, bool RowIserted)
//        {
//            try
//            {
//                oForm.Freeze(true); //행추가여부
//                if (RowIserted == false)
//                {
//                    oDS_PS_AD103L.InsertRecord((oRow));
//                }
//                oMat01.AddRow();
//                oDS_PS_AD103L.Offset = oRow;
//                oDS_PS_AD103L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oMat01.LoadFromDataSource();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// PS_AD103_MTX01
//        /// </summary>
//        private void PS_AD103_MTX01()
//        {
//            string errMessage = string.Empty;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                oForm.Freeze(true);

//            }
//            catch (Exception ex)
//            {
//                if (ProgressBar01 != null)
//                {
//                    ProgressBar01.Stop();
//                }
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
//            }
//        }

//        /// <summary>
//        /// DocEntry 초기화
//        /// </summary>
//        private void PS_AD103_FormClear()
//        {
//            string DocEntry;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_AD103'", "");
//                if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
//                {
//                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//                }
//                else
//                {
//                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// PS_AD103_etBaseForm
//        /// </summary>
//        private void PS_AD103_DeleteAttach(int pRow)
//        {
//            string DeleteFilePath = null;
//            Scripting.FileSystemObject FSO = null;

//            FSO = new Scripting.FileSystemObject();
//            string errMessage = string.Empty;
//            try
//            {
//                oMat01.FlushToDataSource();

//                DeleteFilePath = oDS_PS_AD103L.GetValue("U_AttPath", pRow - 1);
//                //삭제할 첨부파일 경로 저장

//                if (string.IsNullOrEmpty(DeleteFilePath))
//                {

//                    SubMain.Sbo_Application.MessageBox("첨부파일이 없습니다.");

//                }
//                else
//                {

//                    if (SubMain.Sbo_Application.MessageBox("첨부파일을 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
//                    {
//                        FSO.DeleteFile(DeleteFilePath);
//                        //파일 삭제
//                        oDS_PS_AD103L.SetValue("U_AttPath", pRow - 1, "");
//                        //첨부파일 경로 삭제
//                        SubMain.Sbo_Application.MessageBox("파일이 삭제되었습니다.");
//                    }

//                }

//                oMat01.LoadFromDataSource();
//                oMat01.AutoResizeColumns();
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//                }
//            }
//            finally
//            {

//            }
//        }

//        /// <summary>
//        /// PS_AD103_etBaseForm
//        /// </summary>
//        private void PS_AD103_SaveAttach(int pRow)
//        {
//            short i = 0;
//            string sFilePath = null;
//            string sFileName = null;
//            //Dim sQry        As String
//            string SaveFolders = null;
//            object objFile = null;
//            Scripting.FileSystemObject FSO = null;

//            FSO = new Scripting.FileSystemObject();
//            SaveFolders = "\\\\191.1.1.220\\Attach\\PS_AD103";

//            try
//            {
//                sFilePath = My.MyProject.Forms.FileListBoxForm.OpenDialog(ref FileListBoxForm, ref "oxps|*.oxps|xps|*xsp|tif|*.tif", ref "파일선택", ref "C:\\");
//                if (string.IsNullOrEmpty(sFilePath))
//                    return;
//                sFilePath = Strings.Replace(sFilePath, Strings.Chr(0), "");
//                sFileName = Strings.Trim(Strings.Mid(sFilePath, Strings.InStrRev(sFilePath, "\\") + 1, Strings.Len(sFilePath) - 1));
//                //파일명만 추출

//                oMat01.FlushToDataSource();

//                //서버에 기존 파일 체크
//                foreach (object objFile_loopVariable in FSO.GetFolder(SaveFolders).Files)
//                {
//                    objFile = objFile_loopVariable;
//                    //UPGRADE_WARNING: objFile.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (Strings.InStr(Strings.Trim(Strings.UCase(objFile.Name)), Strings.UCase(sFileName)) > 0)
//                    {
//                        if (SubMain.Sbo_Application.MessageBox("동일한 문서번호의 파일이 존재합니다. 교체하시겠습니까?", 2, "Yes", "No") == 1)
//                        {
//                            //기존 파일 삭제
//                            //UPGRADE_WARNING: objFile.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            FileSystem.Kill(SaveFolders + "\\" + objFile.Name);
//                        }
//                        else
//                        {
//                            return;
//                        }
//                    }
//                }

//                oDS_PS_AD103L.SetValue("U_AttPath", pRow - 1, SaveFolders + "\\" + sFileName);
//                //첨부파일 경로 등록

//                oMat01.LoadFromDataSource();
//                oMat01.AutoResizeColumns();

//                FSO.CopyFile(sFilePath, SaveFolders + "\\" + sFileName, true);
//                //파일 복사

//                SubMain.Sbo_Application.MessageBox("업로드 되었습니다.");
//                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//                }
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//                }
//            }
//            finally
//            {

//            }
//        }

//        /// <summary>
//        /// PS_AD103_etBaseForm
//        /// </summary>
//        private void PS_AD103_OpenAttach(int pRow)
//        {
//            string AttachPath = null;
//            string errMessage = string.Empty;

//            try
//            {
//                oMat01.FlushToDataSource();

//                AttachPath = Strings.Trim(oDS_PS_AD103L.GetValue("U_AttPath", pRow - 1));

//                if (string.IsNullOrEmpty(AttachPath))
//                {

//                    SubMain.Sbo_Application.MessageBox("첨부파일이 없습니다.");

//                }
//                else
//                {

//                    ShellExecute(0, "OPEN", AttachPath, Constants.vbNullString, Constants.vbNullString, 5);

//                }
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//                }
//            }
//            finally
//            {

//            }
//        }

//        /// <summary>
//        /// 필수 사항 check
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_AD103_DataValidCheck()
//        {
//            bool ReturnValue = false;
//            int i = 0;
//            string errMessage = string.Empty;
//            string ClickCode = string.Empty;
//            string type = string.Empty;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                    PS_AD103_FormClear();
//                }

//                //사업장 미입력 시
//                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Selected.VALUE))
//                {
//                    errMessage = "사업장이 선택되지 않았습니다.";
//                    throw new Exception();
//                }

//                //라인정보 미입력 시
//                if (oMat01.VisualRowCount == 1)
//                {
//                    errMessage = "라인이 존재하지 않습니다.";
//                    throw new Exception();
//                }

//                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
//                {
//                    //요청자 미입력 시
//                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("ReqCd").Cells.Item(i).Specific.VALUE)))
//                    {
//                        errMessage = "요청자는 필수입니다.";
//                        oMat01.Columns.Item("ReqCd").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        throw new Exception();
//                    }

//                    //기안(요청)일 미입력 시
//                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("ReqDate").Cells.Item(i).Specific.VALUE)))
//                    {
//                        errMessage = "기안(요청)일은 필수입니다.";
//                        oMat01.Columns.Item("ReqDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        throw new Exception();
//                    }

//                    //최종결재일 미입력 시
//                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("LastDate").Cells.Item(i).Specific.VALUE)))
//                    {
//                        errMessage = "최종결재일은 필수입니다.";
//                        oMat01.Columns.Item("LastDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        throw new Exception();
//                    }

//                    //완료희망일 미입력 시
//                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("HopeDate").Cells.Item(i).Specific.VALUE)))
//                    {
//                        errMessage = "완료희망일은 필수입니다.";
//                        oMat01.Columns.Item("HopeDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        throw new Exception();
//                    }

//                    //관련근거(문서번호) 미입력 시
//                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("GWNo").Cells.Item(i).Specific.VALUE)))
//                    {
//                        errMessage = "관련근거(문서번호)는 필수입니다.";
//                        oMat01.Columns.Item("GWNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        throw new Exception();
//                    }
//                }
//                oMat01.FlushToDataSource();
//                oDS_PS_AD103L.RemoveRecord(oDS_PS_AD103L.Size - 1);
//                oMat01.LoadFromDataSource();

//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                    PS_AD103_FormClear();
//                }
//                ReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    if (type == "F")
//                    {
//                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        PSH_Globals.SBO_Application.MessageBox(errMessage);
//                    }
//                    else if (type == "M")
//                    {
//                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        PSH_Globals.SBO_Application.MessageBox(errMessage);
//                    }
//                    else
//                    {
//                        PSH_Globals.SBO_Application.MessageBox(errMessage);
//                    }
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }
//            return ReturnValue;
//        }

//        /// <summary>
//        /// Form Item Event
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">pVal</param>
//        /// <param name="BubbleEvent">Bubble Event</param>
//        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
//                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
//                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
//                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
//                    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
//                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
//                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
//                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
//                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
//                    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
//                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
//                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
//                    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
//                    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
//                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
//                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
//                    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
//                    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
//                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
//                    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
//                    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
//                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
//                    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
//                    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_Drag: //39
//                    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
//                    break;
//            }
//        }

//        /// <summary>
//        /// ITEM_PRESSED 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                            if (PS_AD103_DataValidCheck() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
//                            oFormMode01 = oForm.Mode;
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                        {
//                            if (PS_AD103_DataValidCheck() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.VALUE);
//                            oFormMode01 = oForm.Mode;
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                        {
//                        }
//                    }
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                            if (pVal.ActionSuccess == true)
//                            {
//                                PS_AD103_FormItemEnabled();
//                                PS_AD103_AddMatrixRow(0, true);
//                            }
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                        {
//                            if (pVal.ActionSuccess == true)
//                            {
//                                PS_AD103_FormItemEnabled();
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// KEY_DOWN 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                    if (pVal.ItemUID == "Mat01")
//                    {
//                        if (pVal.ColUID == "ReqCd")
//                        {
//                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ReqCd");
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// GOT_FOCUS 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.ItemUID == "Mat01")
//                {
//                    if (pVal.Row > 0)
//                    {
//                        oLastItemUID01 = pVal.ItemUID;
//                        oLastColUID01 = pVal.ColUID;
//                        oLastColRow01 = pVal.Row;
//                    }
//                }
//                else
//                {
//                    oLastItemUID01 = pVal.ItemUID;
//                    oLastColUID01 = "";
//                    oLastColRow01 = 0;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// COMBO_SELECT 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    if (pVal.ItemUID == "Mat01")
//                    {
//                        if (pVal.ColUID == "Action")
//                        {
//                            if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.VALUE == "S")
//                            {
//                                PS_AD103_SaveAttach(pVal.Row);
//                            }
//                            else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.VALUE == "O")
//                            {
//                                PS_AD103_OpenAttach(pVal.Row);
//                            }
//                            else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.VALUE == "D")
//                            {
//                                PS_AD103_DeleteAttach(pVal.Row);
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// CLICK 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                    if (pVal.ItemUID == "Mat01")
//                    {
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID01 = pVal.ItemUID;
//                            oLastColUID01 = pVal.ColUID;
//                            oLastColRow01 = pVal.Row;

//                            oMat01.SelectRow(pVal.Row, true, false);
//                        }
//                    }
//                    else
//                    {
//                        oLastItemUID01 = pVal.ItemUID;
//                        oLastColUID01 = "";
//                        oLastColRow01 = 0;
//                    }
//                }
//                else if (pVal.BeforeAction == false)
//                {

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// VALIDATE 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                if (pVal.BeforeAction == true)
//                {
//                    if (pVal.ItemChanged == true)
//                    {
//                        if (pVal.ItemUID == "Mat01")
//                        {
//                            if (pVal.ColUID == "ReqCd")
//                            {
//                                oMat01.FlushToDataSource();
//                                oDS_PS_AD103L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
//                                oDS_PS_AD103L.SetValue("U_ReqNm", pVal.Row - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'",""));
                                
//                                if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_AD103L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
//                                {
//                                    PS_AD103_AddMatrixRow(pVal.Row, false);
//                                }
//                                oMat01.LoadFromDataSource();
//                            }
//                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        }
//                        oForm.Update();
//                        oMat01.AutoResizeColumns();
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//                BubbleEvent = false;
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// MATRIX_LOAD 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {

//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    PS_AD103_FormItemEnabled();
//                    PS_AD103_AddMatrixRow(oMat01.VisualRowCount, false);
//                    oMat01.AutoResizeColumns();
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// FORM_UNLOAD 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                    SubMain.Remove_Forms(oFormUniqueID);

//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// EVENT_ROW_DELETE
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                int i = 0;
//                if ((oLastColRow01 > 0))
//                {
//                    if (pVal.BeforeAction == true)
//                    {
//                        //행삭제전 행삭제가능여부검사
//                        if ((PS_AD103_Validate("RowDelete") == false))
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        for (i = 1; i <= oMat01.VisualRowCount; i++)
//                        {
//                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
//                        }
//                        oMat01.FlushToDataSource();
//                        oDS_PS_AD103L.RemoveRecord(oDS_PS_AD103L.Size - 1);
//                        oMat01.LoadFromDataSource();
//                        if (oMat01.RowCount == 0)
//                        {
//                            PS_AD103_AddMatrixRow(0, false);
//                        }
//                        else
//                        {
//                            if (!string.IsNullOrEmpty(oDS_PS_AD103L.GetValue("U_ReqCd", oMat01.RowCount - 1).ToString().Trim()))
//                            {
//                                PS_AD103_AddMatrixRow(oMat01.RowCount, false);
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// FormMenuEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="pVal"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                oForm.Freeze(true);

//                if (pVal.BeforeAction == true)
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284": //취소
//                            break;
//                        case "1286": //닫기
//                            break;
//                        case "1293": //행삭제
//                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
//                            break;
//                        case "1281": //찾기
//                            break;
//                        case "1282": //추가
//                            break;
//                        case "1288": //레코드이동(최초)
//                        case "1289": //레코드이동(이전)
//                        case "1290": //레코드이동(다음)
//                        case "1291": //레코드이동(최종)
//                            break;
//                    }
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284": //취소
//                            break;
//                        case "1286": //닫기
//                            break;
//                        case "1293": //행삭제
//                            Raise_EVENT_ROW_DELETE(ref FormUID, pVal, ref BubbleEvent);
//                            break;
//                        case "1281": //찾기
//                            PS_AD103_FormItemEnabled();
//                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            break;
//                        case "1282": //추가
//                            PS_AD103_FormItemEnabled();
//                            PS_AD103_AddMatrixRow(0, true);
//                            break;
//                        case "1288": //레코드이동(최초)
//                        case "1289": //레코드이동(이전)
//                        case "1290": //레코드이동(다음)
//                        case "1291": //레코드이동(최종)
//                        case "1287": //복제
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// FormDataEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="BusinessObjectInfo"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (BusinessObjectInfo.BeforeAction == true)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
//                            break;
//                    }
//                }
//                else if (BusinessObjectInfo.BeforeAction == false)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// RightClickEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="pVal"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                }

//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID01 = pVal.ItemUID;
//                            oLastColUID01 = pVal.ColUID;
//                            oLastColRow01 = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID01 = pVal.ItemUID;
//                        oLastColUID01 = "";
//                        oLastColRow01 = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//        }
//    }
//}
