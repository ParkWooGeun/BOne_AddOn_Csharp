using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기타레포트조회
    /// </summary>
    internal class PS_QM704 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid oGrid01;
        private SAPbouiCOM.DataTable oDS_PS_QM704H;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_QM704L;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM704.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_QM704_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_QM704");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                
                PS_QM704_CreateItems();
                PS_QM704_EnableMenus();
                PS_QM704_ComboBox_Setting();
                PS_QM704_AddMatrixRow(0, true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PS_QM704_CreateItems()
        {
            try
            {
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_QM704H");
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_QM704H");
                oDS_PS_QM704H = oForm.DataSources.DataTables.Item("PS_QM704H");

                oDS_PS_QM704L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("oMat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //기간
                oForm.DataSources.UserDataSources.Add("DocDatefr", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDatefr").Specific.DataBind.SetBound(true, "", "DocDatefr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("DocDateto", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateto").Specific.DataBind.SetBound(true, "", "DocDateto");
                oForm.DataSources.UserDataSources.Item("DocDateto").Value = DateTime.Now.ToString("yyyyMMdd");
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
        /// 매트릭스 행 추가
        /// PH_PY035_Add_MatrixRow
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// </summary>
        private void PS_QM704_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_QM704L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_QM704L.Offset = oRow;
                oDS_PS_QM704L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM702H_Add_MatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// EnableMenus 메뉴설정
        /// </summary>
        private void PS_QM704_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                dataHelpClass.SetEnableMenus(oForm, false, false, false, false, false, true, false, false, false, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_QM704_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList((oForm.Items.Item("BPLId").Specific), "SELECT BPLId, BPLName From [OBPL] order by 1", "", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("InOut").Specific.ValidValues.Add("I", "자체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("O", "외주");
                oForm.Items.Item("InOut").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("CheckYN").Specific.ValidValues.Add("1", "승인");
                oForm.Items.Item("CheckYN").Specific.ValidValues.Add("2", "반려");
                oForm.Items.Item("CheckYN").Specific.ValidValues.Add("3", "결재대기");
                oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //Action(Matrix)
                sQry = "  SELECT      U_Minor, ";
                sQry += "             U_CdName ";
                sQry += " FROM        [@PS_SY001L] ";
                sQry += " WHERE       Code = 'A009'";
                sQry += "             AND ISNULL(U_UseYN, 'Y') = 'Y'";
                sQry += " ORDER BY    U_Seq";

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Action"), sQry, "", "");
                
                oMat01.Columns.Item("endsoL").ValidValues.Add("O", "완료");
                oMat01.Columns.Item("endsoL").ValidValues.Add("X", "미완료");

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
        /// PS_QM704_etBaseForm
        /// </summary>
        private void PS_QM704_DeleteAttach(int pRow)
        {
            string DeleteFilePath;
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                DeleteFilePath = oDS_PS_QM704L.GetValue("U_ColRgL01", pRow - 1); //삭제할 첨부파일 경로 저장

                if (string.IsNullOrEmpty(DeleteFilePath))
                {
                    errMessage = "첨부파일이 없습니다.";
                }
                else
                {
                    if (PSH_Globals.SBO_Application.MessageBox("첨부파일을 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                    {
                        System.IO.File.Delete(DeleteFilePath);
                        oDS_PS_QM704L.SetValue("U_ColRgL01", pRow - 1, ""); //첨부파일 경로 삭제
                        PSH_Globals.SBO_Application.MessageBox("파일이 삭제되었습니다.");
                    }
                    if (oDS_PS_QM704L.GetValue("U_ColReg01", pRow - 1) == "외주")
                    {
                        sQry = "UPDATE [@PS_QM701H] SET U_AttPath ='' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                    else
                    {
                        sQry = "UPDATE [@PS_QM703H] SET U_AttPath ='' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_QM704_etBaseForm
        /// </summary>
        private void PS_QM704_SaveAttach(int pRow)
        {
            string sFileFullPath;
            string sFilePath;
            string sFileName;
            string SaveFolders;
            string sourceFile;
            string targetFile;
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sFileFullPath = PS_QM704_OpenFileSelectDialog();//OpenFileDialog를 쓰레드로 실행

                SaveFolders = "\\\\191.1.1.220\\Attach\\PS_QM704";
                sFileName = System.IO.Path.GetFileName(sFileFullPath); //파일명
                sFilePath = System.IO.Path.GetDirectoryName(sFileFullPath); //파일명을 제외한 전체 경로

                sourceFile = System.IO.Path.Combine(sFilePath, sFileName);
                targetFile = System.IO.Path.Combine(SaveFolders, sFileName);
                oMat01.FlushToDataSource();

                if (System.IO.File.Exists(targetFile)) //서버에 기존파일이 존재하는지 체크
                {
                    if (PSH_Globals.SBO_Application.MessageBox("동일한 문서번호의 파일이 존재합니다. 교체하시겠습니까?", 2, "Yes", "No") == 1)
                    {
                        System.IO.File.Delete(targetFile); //삭제
                    }
                    else
                    {
                        return;
                    }
                }
                oDS_PS_QM704L.SetValue("U_ColRgL01", pRow - 1, SaveFolders + "\\" + sFileName); //첨부파일 경로 등록

                if(oDS_PS_QM704L.GetValue("U_ColReg01", pRow - 1).ToString().Trim() =="외주")
                {
                    sQry = "UPDATE [@PS_QM701H] SET U_AttPath ='" + oDS_PS_QM704L.GetValue("U_ColRgL01", pRow - 1) + "' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).Trim() + "'";
                    oRecordSet01.DoQuery(sQry);
                }
                else
                {
                    sQry = "UPDATE [@PS_QM703H] SET U_AttPath ='" +( SaveFolders + "\\" + sFileName) + "' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).Trim() + "'";
                    oRecordSet01.DoQuery(sQry);
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                System.IO.File.Copy(sourceFile, targetFile, true); //파일 복사 (여기서 오류발생)
                PSH_Globals.SBO_Application.MessageBox("업로드 되었습니다.");
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// OpenFileSelectDialog 호출(쓰레드를 이용하여 비동기화)
        /// OLE 호출을 수행하려면 현재 스레드를 STA(단일 스레드 아파트) 모드로 설정해야 합니다.
        /// </summary>
        [STAThread]
        private string PS_QM704_OpenFileSelectDialog()
        {
            string returnFileName = string.Empty;

            var thread = new System.Threading.Thread(() =>
            {
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1; //FilterIndex는 1부터 시작
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    returnFileName = openFileDialog.FileName;
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            return returnFileName;
        }

        /// <summary>
        /// PS_QM704_etBaseForm
        /// </summary>
        private void PS_QM704_OpenAttach(int pRow)
        {
            string AttachPath;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                AttachPath = oDS_PS_QM704L.GetValue("U_ColRgL01", pRow - 1).ToString().Trim();
                if (string.IsNullOrEmpty(AttachPath))
                {
                    PSH_Globals.SBO_Application.MessageBox("첨부파일이 없습니다.");
                }
                else
                {
                    System.Diagnostics.ProcessStartInfo process = new System.Diagnostics.ProcessStartInfo(AttachPath);
                    process.UseShellExecute = true;
                    process.Verb = "open";

                    System.Diagnostics.Process.Start(process);
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
        }

        /// <summary>
        /// PS_QM704_Update
        /// </summary>
        private void PS_QM704_Update(int pRow, int cnt)
        {
          
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oMat01.FlushToDataSource();

                if (cnt == 1)
                {
                    if (oDS_PS_QM704L.GetValue("U_ColReg01", pRow - 1).ToString().Trim() == "외주")
                    {
                        sQry = "UPDATE [@PS_QM701H] SET U_endsoL ='O' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                    else
                    {
                        sQry = "UPDATE [@PS_QM703H] SET U_endsoL ='O' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                else
                {
                    if (oDS_PS_QM704L.GetValue("U_ColReg01", pRow - 1).ToString().Trim() == "외주")
                    {
                        sQry = "UPDATE [@PS_QM701H] SET U_endsoL ='X' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                    else
                    {
                        sQry = "UPDATE [@PS_QM703H] SET U_endsoL ='X' WHERE DocEntry ='" + oDS_PS_QM704L.GetValue("U_ColReg02", pRow - 1).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_QM704_MTX01
        /// </summary>
        private void PS_QM704_MTX01()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            string Query01 = string.Empty;
            string BPLId;
            string DocDateFr;
            string DocDateTo;
            string ChkYN;
            string Gobun;
            string errCode = string.Empty;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                oForm.Freeze(true);
                
                BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
                DocDateFr = oForm.Items.Item("DocDatefr").Specific.Value.ToString().Trim();
                DocDateTo = oForm.Items.Item("DocDateto").Specific.Value.ToString().Trim();
                Gobun = oForm.Items.Item("InOut").Specific.Value.ToString().Trim();
                ChkYN = oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim();

                if(Gobun == "O")
                {
                    Query01 = "EXEC PS_QM704_01 '" + BPLId + "','" + DocDateFr + "','" + DocDateTo + "','" + ChkYN + "'";
                }
                else
                {
                    Query01 = "EXEC PS_QM704_02 '" + BPLId + "','" + DocDateFr + "','" + DocDateTo + "','" + ChkYN + "'";
                }
                oGrid01.DataTable.Clear();

                oDS_PS_QM704H.ExecuteQuery(Query01);

                if (oGrid01.Rows.Count == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                oGrid01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                oForm.Freeze(false);
                oForm.Update();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_QM704_LoadData(int p_DocEntry, string p_Gobun)
        {
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oDS_PS_QM704L.Clear(); //추가

                sQry = "EXEC [PS_QM704_03] '" + p_DocEntry + "','" + p_Gobun + "'";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "취소된 문건입니다.";
                    throw new Exception();
                }
                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_QM704L.Size)
                    {
                        oDS_PS_QM704L.InsertRecord((i));
                    }
                    oMat01.AddRow();
                    oDS_PS_QM704L.Offset = i;
                    oDS_PS_QM704L.SetValue("U_LineNum", i, Convert.ToString(1));  // 순번
                    oDS_PS_QM704L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Gobun").Value.ToString().Trim());  // 관리번호
                    oDS_PS_QM704L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());    // 시작일자
                    oDS_PS_QM704L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());    // 시작시간
                    oDS_PS_QM704L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());    // 종료일자
                    oDS_PS_QM704L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_WorkNum").Value.ToString().Trim());    // 종료시간
                    oDS_PS_QM704L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_ITemName").Value.ToString().Trim());    // 사용차량
                    oDS_PS_QM704L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_ItemSpec").Value.ToString().Trim());    // 목적지
                    oDS_PS_QM704L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_WorkCode").Value.ToString().Trim());   // 신청차사번
                    oDS_PS_QM704L.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_WorkName").Value.ToString().Trim());   // 신청자명
                    oDS_PS_QM704L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("u_mstcod").Value.ToString().Trim());     // 동승자명
                    oDS_PS_QM704L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("u_mstnam").Value.ToString().Trim());  // 주행전Km
                    oDS_PS_QM704L.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("U_verdict").Value.ToString().Trim());  // 주행전Km
                    oDS_PS_QM704L.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("u_chkYN").Value.ToString().Trim());    // 주행후Km
                    oDS_PS_QM704L.SetValue("U_ColRgL01", i, oRecordSet01.Fields.Item("U_AttPath").Value.ToString().Trim());   // 등록구분
                    oDS_PS_QM704L.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("U_Action").Value.ToString().Trim());   // 등록구분
                    oDS_PS_QM704L.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("U_endsoL").Value.ToString().Trim());   // 등록구분
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
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PS_QM704_MTX01:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
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

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);+-
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
                    PS_QM704_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            PS_QM704_MTX01();
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
            int DOCENTRY;
            string GOBUN;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        DOCENTRY = oDS_PS_QM704H.Columns.Item("문서번호").Cells.Item(pVal.Row).Value;
                        GOBUN = oForm.Items.Item("InOut").Specific.Value.ToString().Trim();
                        PS_QM704_LoadData(DOCENTRY, GOBUN);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    if (pVal.ItemUID == "oMat01")
                    {
                        if (pVal.ColUID == "Action" && !string.IsNullOrEmpty(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value))
                        {
                            if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "S")
                            {
                                PS_QM704_SaveAttach(pVal.Row);
                            }
                            else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "O")
                            {
                                PS_QM704_OpenAttach(pVal.Row);
                            }
                            else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "D")
                            {
                                PS_QM704_DeleteAttach(pVal.Row);
                            }
                        }
                        if (pVal.ColUID == "endsoL" && !string.IsNullOrEmpty(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value))
                        {
                            if (oMat01.Columns.Item("endsoL").Cells.Item(pVal.Row).Specific.Value == "O")
                            {
                                PS_QM704_Update(pVal.Row,1);
                            }
                            else if (oMat01.Columns.Item("endsoL").Cells.Item(pVal.Row).Specific.Value == "X")
                            {
                                PS_QM704_Update(pVal.Row,2);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_QM704_FormResize()
        {
            try
            {
                oForm.Items.Item("Grid01").Top = 80;
                oForm.Items.Item("Grid01").Left = 7;
                oForm.Items.Item("Grid01").Height = (oForm.Height / 3) *2  - 20;
                oForm.Items.Item("Grid01").Width = (oForm.Width);

                oForm.Items.Item("Item_0").Top = (oForm.Height / 3) *2 + 80;
                oForm.Items.Item("Item_0").Left = 7;
                oForm.Items.Item("Item_0").Height = 15;
                oForm.Items.Item("Item_0").Width = 80;


                oForm.Items.Item("oMat01").Top = (oForm.Height / 3) * 2 + 100;
                oForm.Items.Item("oMat01").Left = 7;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    PS_QM704_FormResize();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM704H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM704L);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
