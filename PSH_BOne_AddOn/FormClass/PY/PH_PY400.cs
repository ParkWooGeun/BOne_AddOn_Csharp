using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연말정산대상자등록
    /// </summary>
    internal class PH_PY400 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY400B; //등록라인
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY400.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY400_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY400");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry"

                oForm.Freeze(true);
                PH_PY400_CreateItems();
                PH_PY400_ComboBox_Setting();
                PH_PY400_EnableMenus();
                PH_PY400_SetDocument(oFormDocEntry);
                PH_PY400_FormResize();
                PH_PY400_Add_MatrixRow(0, true);
                PH_PY400_LoadCaption();
                PH_PY400_FormReset();
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
        private void PH_PY400_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY400B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

                //정산년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.DataBind.SetBound(true, "", "YYYY");
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
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY400_ComboBox_Setting()
        {
            try
            {
                oForm.Freeze(true);
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
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY400_EnableMenus()
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY400_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY400_FormItemEnabled();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY400_FormItemEnabled()
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY400_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PH_PY400_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PH_PY400B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY400B.Offset = oRow;
                oDS_PH_PY400B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PH_PY400B.SetValue("U_ColReg01", oRow, "Y");

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PH_PY400_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "저장";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
        /// 화면 초기화
        /// </summary>
        private void PH_PY400_FormReset()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                string User_BPLID = dataHelpClass.User_BPLID();

                //헤더 초기화
                oForm.DataSources.UserDataSources.Item("CLTCOD").Value = User_BPLID; //사업장
                oForm.DataSources.UserDataSources.Item("YYYY").Value = Convert.ToString(DateTime.Now.Year - 1); //정산년도

                //라인 초기화
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                PH_PY400_Add_MatrixRow(0, true);
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
        /// 데이터 조회
        /// </summary>
        private void PH_PY400_MTX01()
        {
            int i;
            string sQry;
            short ErrNum = 0;

            string saup;
            string yyyy;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();  //사업장
                yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();    //년도

                sQry = "   SELECT a.sabun  AS [sabun], ";  
                sQry += " 		  b.U_FullName AS [FullName], ";
                sQry += "         b.U_TeamCode AS [TeamCode], ";
                sQry += "         c.U_CodeNm AS [TeamName], ";
                sQry += "         isnull(b.U_RspCode,'') AS [RspCode], ";
                sQry += "         isnull(d.U_CodeNm,'') AS [RspName], ";
                sQry += "         Convert(char(10), b.U_TermDate, 23) AS [TermDate], ";
                sQry += "         a.ChkYN AS [ChkYN], ";
                sQry += " 		  c.U_Seq, ";
                sQry += "         d.U_Seq  ";
                sQry += "    FROM [p_seoytarget] a INNER JOIN [@PH_PY001A] AS b ON a.saup = b.U_CLTCOD AND a.sabun = b.Code ";
                sQry += "                          LEFT JOIN  [@PS_HR200L] AS c ON b.U_TeamCode = c.U_Code AND c.Code = '1' ";
                sQry += "                          LEFT JOIN  [@PS_HR200L] AS d ON b.U_RspCode = d.U_Code  AND d.Code = '2' ";
                sQry += "  WHERE a.saup = '" + saup + "' ";
                sQry += "    AND a.yyyy = '" + yyyy + "' ";
                sQry += " ORDER BY c.U_Seq, ";
                sQry += "          d.U_Seq, ";
                sQry += " 		   a.sabun ";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY400B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 1;

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    PH_PY400_Add_MatrixRow(0, true);
                    PH_PY400_LoadCaption();

                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY400B.Size)
                    {
                        oDS_PH_PY400B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY400B.Offset = i;

                    oDS_PH_PY400B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY400B.SetValue("U_ColReg01", i, "Y");                                                        //선택Y
                    oDS_PH_PY400B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("sabun").Value.ToString().Trim());    //사번
                    oDS_PH_PY400B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("FullName").Value.ToString().Trim()); //성명
                    oDS_PH_PY400B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim()); //부서
                    oDS_PH_PY400B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim()); //부서명
                    oDS_PH_PY400B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());  //담당
                    oDS_PH_PY400B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());  //담당명
                    oDS_PH_PY400B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("TermDate").Value.ToString().Trim()); //퇴직일자
                    oDS_PH_PY400B.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("ChkYN").Value.ToString().Trim());    //계산여부

                    oRecordSet.MoveNext();
                }

                PH_PY400_Add_MatrixRow(oMat01.VisualRowCount, false);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 데이터 저장(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PH_PY400_AddData()
        {
            bool returnValue = false;

            short i = 0;
            string sQry = string.Empty;
            string saup = string.Empty;
            string yyyy = string.Empty;
            string sabun = string.Empty;
            string ChkYN = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (oDS_PH_PY400B.GetValue("U_ColReg01", i).ToString().Trim() == "Y" && (!string.IsNullOrEmpty(oDS_PH_PY400B.GetValue("U_ColReg02", i).ToString().Trim()))) //선택'Y'시 AND 사번있을시
                    {
                        saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();   //사업장
                        yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();     //년도
                        sabun = oDS_PH_PY400B.GetValue("U_ColReg02", i).ToString().Trim();    //사번
                        ChkYN = oDS_PH_PY400B.GetValue("U_ColReg09", i).ToString().Trim();    //정산계산여부

                        sQry = " Select Count(*) From [p_seoytarget] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item(0).Value > 0)
                        {
                            // 갱신
                            sQry = "Update [p_seoytarget] set ";
                            sQry += "ChkYN = '" + ChkYN + "'";
                            sQry += " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                        else
                        {
                            // 신규
                            sQry = "INSERT INTO [p_seoytarget]";
                            sQry += " (";
                            sQry += "saup,";
                            sQry += "yyyy,";
                            sQry += "sabun,";
                            sQry += "ChkYN )";
                            sQry += " VALUES(";
                            sQry += "'" + saup + "',";
                            sQry += "'" + yyyy + "',";
                            sQry += "'" + sabun + "',";
                            sQry += "'" + ChkYN + "')";
                            oRecordSet.DoQuery(sQry);
                        }
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 삭제
        /// </summary>
        private void PH_PY400_DeleteData()
        {
            short i;
            short ErrNum = 0;
            string sQry;
            string saup;
            string yyyy;
            string sabun;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.SBO_Application.MessageBox("선택한 자료를 삭제 하시겠습니까?.", 1, "예", "아니오") == 1)
                {
                    oMat01.FlushToDataSource();

                    saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();   //사업장
                    yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();     //년도

                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (oDS_PH_PY400B.GetValue("U_ColReg01", i).ToString().Trim() == "Y" && (!string.IsNullOrEmpty(oDS_PH_PY400B.GetValue("U_ColReg02", i).ToString().Trim()))) //선택'Y'시 AND 사번있을시
                        {
                            sabun = oDS_PH_PY400B.GetValue("U_ColReg02", i).ToString().Trim();    //사번

                            sQry = "Delete From [p_seoytarget] Where saup = '" + saup + "' AND  yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                    }
                }
                else
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("삭제가 취소 되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메트릭스 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PH_PY400_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i = 0;
            short ErrNum = 0;

            try
            {
                oMat01.FlushToDataSource();

                for (i = 0; i <= oMat01.VisualRowCount - 2; i++) //마지막 빈행 제외를 위해 2를 뺌
                {
                    if (oDS_PH_PY400B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {

                        if (string.IsNullOrEmpty(oDS_PH_PY400B.GetValue("U_ColReg02", i).ToString().Trim())) //사번
                        {
                            ErrNum = 1;
                            throw new Exception();
                        }
                    }
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + 1 + "번 라인의 사번이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY400_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;

            string TeamCode; //대상자부서
            string TeamName;
            string RspCode; //대상자담당
            string RspName;
            string FullName; //성명
            string TermDate; //퇴직일자
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "Mat01":

                        oMat01.FlushToDataSource();

                        if (oCol == "sabun")
                        {
                            oDS_PH_PY400B.SetValue("U_ColReg02", oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim()); //사번

                            //대상자의 인사마스터에서 소속 조회
                            sQry = "  SELECT  T0.U_TeamCode AS [TeamCode], "; //부서코드
                            sQry += "         T1.U_CodeNm AS [TeamName], "; //부서명
                            sQry += "         T0.U_RspCode AS [RspCode], "; //담당코드
                            sQry += "         T2.U_CodeNm AS [RspName], "; //담당명
                            sQry += "         T0.U_FullName AS [FullName], "; //성명
                            sQry += "         T0.U_TermDate AS [TermDate] "; //퇴직일자
                            sQry += " FROM    [@PH_PY001A] AS T0 ";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T1";
                            sQry += "             ON T0.U_TeamCode = T1.U_Code";
                            sQry += "             AND T1.Code = '1'";
                            sQry += "         LEFT JOIN";
                            sQry += "         [@PS_HR200L] AS T2";
                            sQry += "             ON T0.U_RspCode = T2.U_Code";
                            sQry += "             AND T2.Code = '2'";
                            sQry += " WHERE   T0.Code = '" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);

                            TeamCode = oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim();
                            TeamName = oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim();
                            RspCode = oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim();
                            RspName = oRecordSet01.Fields.Item("RspName").Value.ToString().Trim();
                            FullName = oRecordSet01.Fields.Item("FullName").Value.ToString().Trim();
                            TermDate = oRecordSet01.Fields.Item("TermDate").Value.ToString().Trim();

                            oDS_PH_PY400B.SetValue("U_ColReg03", oRow - 1, FullName);
                            oDS_PH_PY400B.SetValue("U_ColReg04", oRow - 1, TeamCode);
                            oDS_PH_PY400B.SetValue("U_ColReg05", oRow - 1, TeamName);
                            oDS_PH_PY400B.SetValue("U_ColReg06", oRow - 1, RspCode); 
                            oDS_PH_PY400B.SetValue("U_ColReg07", oRow - 1, RspName); 
                            oDS_PH_PY400B.SetValue("U_ColReg08", oRow - 1, Convert.ToDateTime(TermDate).ToString("yyyy-MM-dd"));
                            
                            //행 추가
                            if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PH_PY400B.GetValue("U_ColReg02", oRow - 1).ToString().Trim()))
                            {
                                PH_PY400_Add_MatrixRow(oRow, false);
                            }

                         //   oMat01.Columns.Item("JIGNAM").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        }

                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Matrix 체크박스 전체 선택
        /// </summary>
        private void PH_PY400_CheckAll()
        {
            string CheckType;
            short loopCount;

            CheckType = "Y";

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY400B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY400B.Offset = loopCount;
                    if (CheckType == "N")
                    {
                        oDS_PH_PY400B.SetValue("U_ColReg01", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY400B.SetValue("U_ColReg01", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
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
        /// 정산대상자 불러오기
        /// </summary>
        private void PH_PY400_Upload()
        {
            int oRow;
            short ErrNum = 0;
            string CLTCOD;
            string YYYY;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();  //사업장
                YYYY = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim(); //년도
                if (string.IsNullOrEmpty(YYYY))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (PSH_Globals.SBO_Application.MessageBox("데이터가 존재하면 '조회-->삭제' 후 실행하세요", 1, "예", "아니오") == 1)
                {
                    oMat01.Clear();
                    oDS_PH_PY400B.Clear();

                    sQry = "   SELECT Distinct ";
                    sQry += "         a.MSTCOD AS [MSTCOD], ";
                    sQry += " 		  b.U_FullName AS [FullName], ";
                    sQry += "         b.U_TeamCode AS [TeamCode], ";
                    sQry += "         c.U_CodeNm AS [TeamName], ";
                    sQry += "         isnull(b.U_RspCode,'') AS [RspCode], ";
                    sQry += "         isnull(d.U_CodeNm,'') AS [RspName], ";
                    sQry += "         Convert(char(10), b.U_TermDate, 23) AS [TermDate], ";
                    sQry += " 		  c.U_Seq, ";
                    sQry += "         d.U_Seq ";
                    sQry += "    FROM (   SELECT Distinct ";
                    sQry += " 				     a.U_MSTCOD AS [MSTCOD] ";
                    sQry += " 			    FROM [@PH_PY112A] a  ";
                    sQry += " 			   WHERE a.U_CLTCOD = '" + CLTCOD + "' ";
                    sQry += " 			     AND a.U_YM Between '" + YYYY + "' + '01' AND '" + YYYY + "' + '12' ";
                    sQry += " 			  Union All ";
                    sQry += "             SELECT a.CODE AS [MSTCOD] ";
                    sQry += " 			    FROM [@PH_PY001A] a  ";
                    sQry += " 			   WHERE a.U_CLTCOD = '" + CLTCOD + "' ";
                    sQry += " 			     AND Convert(char(4), a.U_StartDat, 112) = '" + YYYY + "' ";  //당해입사자중
                    sQry += " 			     AND a.U_JIGTYP <> '06' ";    //파견제외
                    sQry += "          ) a   INNER JOIN [@PH_PY001A] AS b ON b.U_CLTCOD = '" + CLTCOD + "' AND a.MSTCOD = b.Code ";
                    sQry += "                LEFT JOIN  [@PS_HR200L] AS c ON b.U_TeamCode = c.U_Code AND c.Code = '1' ";
                    sQry += "                LEFT JOIN  [@PS_HR200L] AS d ON b.U_RspCode = d.U_Code  AND d.Code = '2' ";
                    sQry += " ORDER BY c.U_Seq, ";
                    sQry += "          d.U_Seq, ";
                    sQry += " 		   a.MSTCOD ";

                    oRecordSet.DoQuery(sQry);
                }
                else
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                oRow = 0;

                while (!oRecordSet.EoF)
                {
                    oDS_PH_PY400B.InsertRecord(oRow);
                    oDS_PH_PY400B.Offset = oRow;
                    oDS_PH_PY400B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));                                      //라인번호  
                    oDS_PH_PY400B.SetValue("U_ColReg01", oRow, "Y");                                                            //선택
                    oDS_PH_PY400B.SetValue("U_ColReg02", oRow, oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim());       //사원번호
                    oDS_PH_PY400B.SetValue("U_ColReg03", oRow, oRecordSet.Fields.Item("FullName").Value.ToString().Trim());     //성명
                    oDS_PH_PY400B.SetValue("U_ColReg04", oRow, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());     //부서
                    oDS_PH_PY400B.SetValue("U_ColReg05", oRow, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim());     //부서명
                    oDS_PH_PY400B.SetValue("U_ColReg06", oRow, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());      //담당
                    oDS_PH_PY400B.SetValue("U_ColReg07", oRow, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());      //담당명
                    oDS_PH_PY400B.SetValue("U_ColReg08", oRow, oRecordSet.Fields.Item("TermDate").Value.ToString().Trim());     //퇴직일자
                    if (string.IsNullOrEmpty(oRecordSet.Fields.Item("TermDate").Value.ToString().Trim()))
                    {
                        oDS_PH_PY400B.SetValue("U_ColReg09", oRow, "Y");                                                            //정산계산Y 정상근무자
                    }
                    else
                    {
                        oDS_PH_PY400B.SetValue("U_ColReg09", oRow, "N");                                                            //정산계산N 퇴직자
                    }
                    oRow += 1;
                    oRecordSet.MoveNext();
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("작업이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("정산년도는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("생성할 대상자가 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("정산대상자불러오기가 취소 되었습니다..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

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

                ////case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                ////    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                ////    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                ////    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                ////    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                ////    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    //    break;

                    //    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    //    break;
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
                    if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY400_MatrixSpaceLineDel() == false) //매트릭스 필수자료 체크
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PH_PY400_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY400_LoadCaption();
                            PH_PY400_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                        PH_PY400_LoadCaption();
                        PH_PY400_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {
                        PH_PY400_DeleteData();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE; //fm_VIEW_MODE

                        PH_PY400_LoadCaption();
                        PH_PY400_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnSel") //전체선택
                    {
                        PH_PY400_CheckAll();
                    }
                    else if (pVal.ItemUID == "BtnLoad") //대상자불러오기
                    {
                        PH_PY400_Upload();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY400")
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    }
                    else
                    {
                        PH_PY400_FlushToItemValue(pVal.ItemUID, 0, "");
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "sabun")
                            {
                                PH_PY400_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                        }
                        else
                        {
                            PH_PY400_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY400_FormItemEnabled();
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY400B);
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
                    PH_PY400_FormResize();
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
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY400_FormReset();
                            PH_PY400_LoadCaption();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
