//using System;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 
//    /// </summary>
//    internal class PS_FX010 : PSH_BaseClass
//    {
//        private string oFormUniqueID;
//        private SAPbouiCOM.Matrix oMat01;
//        private SAPbouiCOM.DBDataSource oDS_PS_FX010H; //등록헤더
//        private SAPbouiCOM.DBDataSource oDS_PS_FX010L; //등록라인

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
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX010.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PS_FX010_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PS_FX010");

//                string strXml = null;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                oForm.Freeze(true);
//                PS_FX010_CreateItems();
//                PS_FX010_ComboBox_Setting();
//                PS_FX010_CF_ChooseFromList();
//                PS_FX010_EnableMenus();
//                PS_FX010_SetDocument(oFromDocEntry01);
//                PS_FX010_FormResize();

//                PS_FX010_AddMatrixRow(0, true);
//                PS_FX010_LoadCaption();
//                PS_FX010_FormItemEnabled();
//                PS_FX010_FormClear();

//                oForm.EnableMenu(("1283"), false); // 삭제
//                oForm.EnableMenu(("1286"), false); // 닫기
//                oForm.EnableMenu(("1287"), false); // 복제
//                oForm.EnableMenu(("1285"), false); // 복원
//                oForm.EnableMenu(("1284"), true); // 취소
//                oForm.EnableMenu(("1293"), false); // 행삭제
//                oForm.EnableMenu(("1281"), false);
//                oForm.EnableMenu(("1282"), true);

               
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
//        private void PS_FX010_CreateItems()
//        {
//            string oQuery01 = null;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                oDS_PS_FX010H = oForm.DataSources.DBDataSources.Item("@PS_FX010H");
//                oDS_PS_FX010L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

//                //// 메트릭스 개체 할당
//                oMat01 = oForm.Items.Item("Mat01").Specific;
//                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat01.AutoResizeColumns();

//                //사업장_S
//                oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");
//                //사업장_E


//                //이력구분_S
//                oForm.DataSources.UserDataSources.Add("SHisType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
//                oForm.Items.Item("SHisType").Specific.DataBind.SetBound(true, "", "SHisType");
//                //이력구분_E

//                //자산코드(검색용)_S
//                oForm.DataSources.UserDataSources.Add("CFixCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("CFixCode").Specific.DataBind.SetBound(true, "", "CFixCode");
//                //자산코드(검색용)_E

//                //이력일자_From_S
//                oForm.DataSources.UserDataSources.Add("DocDateF", SAPbouiCOM.BoDataType.dt_DATE);
//                oForm.Items.Item("DocDateF").Specific.DataBind.SetBound(true, "", "DocDateF");
//                //이력일자_From_E

//                //이력일자_To_S
//                oForm.DataSources.UserDataSources.Add("DocDateT", SAPbouiCOM.BoDataType.dt_DATE);
//                oForm.Items.Item("DocDateT").Specific.DataBind.SetBound(true, "", "DocDateT");
//                //이력일자_To_E

//                oForm.Items.Item("DocDateF").Specific.VALUE = DateTime.Now.ToString("yyyyMM") + "01"; 
//                oForm.Items.Item("DocDateT").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// Combobox 설정
//        /// </summary>
//        private void PS_FX010_ComboBox_Setting()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
//                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

//                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

//                oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

//                //이력구분
//                oForm.Items.Item("HisType").Specific.ValidValues.Add("%", "전체");
//                dataHelpClass.Set_ComboList(oForm.Items.Item("HisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", false, false);
//                oForm.Items.Item("HisType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                //자산분류
//                oForm.Items.Item("ClasCode").Specific.ValidValues.Add("%", "전체");
//                dataHelpClass.Set_ComboList(oForm.Items.Item("ClasCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", false, false);
//                oForm.Items.Item("ClasCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                //이력구분(조회조건)
//                oForm.Items.Item("SHisType").Specific.ValidValues.Add("%", "전체");
//                dataHelpClass.Set_ComboList(oForm.Items.Item("SHisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", false, false);
//                oForm.Items.Item("SHisType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                oForm.Items.Item("AmtYN").Specific.ValidValues.Add("N", "N");
//                oForm.Items.Item("AmtYN").Specific.ValidValues.Add("Y", "Y");
//                oForm.Items.Item("AmtYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                //사업장(매트릭스)
//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId","" ,"");

//                //이력구분(매트릭스)
//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("HisType"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", "");

//                //자산구분(매트릭스)
//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ClasCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", "");
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// Initialization
//        /// </summary>
//        private void PS_FX010_Initialization()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {

//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }


//        /// <summary>
//        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
//        /// </summary>
//        private void PS_FX010_LoadCaption()
//        {
//            try
//            {
//                oForm.Freeze(true);

//                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                {
//                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
//                    oForm.Items.Item("BtnDelete").Enabled = false;
//                }
//                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                {
//                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
//                    oForm.Items.Item("BtnDelete").Enabled = true;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }

//        }
//        /// <summary>
//        /// HeaderSpaceLineDel
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_FX010_HeaderSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            string errMessage = string.Empty;
//            short ErrNum = 0;

//            try
//            {

//                if(oForm.Items.Item("HisType").Specific.VALUE.ToString().Trim() == "%")
//                {
//                    errMessage = "";
//                    throw new Exception();
//                }
//                else if (oForm.Items.Item("ClasCode").Specific.VALUE.ToString().Trim() == "%")
//                {
//                    errMessage = "자산구분은 필수사항입니다. 확인하세요.";
//                    throw new Exception();
//                }
//                else if (string.IsNullOrEmpty(oForm.Items.Item("FixCode").Specific.VALUE.ToString().Trim()))
//                {
//                    errMessage = "자산코드는 필수사항입니다. 확인하세요.";
//                    throw new Exception();
//                }
//                else if (string.IsNullOrEmpty(oForm.Items.Item("FixName").Specific.VALUE.ToString().Trim()))
//                {
//                    errMessage = "자산명은 필수사항입니다. 확인하세요.";
//                    throw new Exception();
//                }
//                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim()))
//                {
//                    errMessage = "이력일자는 필수사항입니다. 확인하세요.";
//                    throw new Exception();
//                }
//                else if (oForm.Items.Item("AmtYN").Specific.VALUE.ToString().Trim() == "Y")
//                {
//                    errMessage = "자본적지출시 금액은 필수사항입니다. 확인하세요.";
//                    throw new Exception();
//                }
//                else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.VALUE.ToString().Trim()))
//                {
//                    errMessage = "매각등록시 거래처는 필수사항입니다. 확인하세요.";
//                    throw new Exception();
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
//            return functionReturnValue;
//        }

//        /// <summary>
//        /// MatrixSpaceLineDel
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_FX010_MatrixSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            int i;
//            string errMessage = string.Empty;

//            try
//            {

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
//            return functionReturnValue;
//        }

//        /// <summary>
//        /// Delete_EmptyRow
//        /// </summary>
//        private void PS_FX010_Delete_EmptyRow()
//        {
//            int i;
//            string errMessage = string.Empty;

//            try
//            {

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
//        }

//        /// <summary>
//        /// ChooseFromList
//        /// </summary>
//        private void PS_FX010_CF_ChooseFromList()
//        {
//            try
//            {
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
//        private bool PS_FX010_Validate(string ValidateType)
//        {
//            bool returnValue = false;
//            int i;
//            int j;
//            string query01;
//            bool Exist;
//            string errCode = string.Empty;
//            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                returnValue = true;
//            }
//            catch (Exception ex)
//            {
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
//            }
//            return returnValue;
//        }

//        /// <summary>
//        /// PS_FX010_AddMatrixRow
//        /// </summary>
//        /// <param name="oRow">행 번호</param>
//        /// <param name="RowIserted">행 추가 여부</param>
//        private void PS_FX010_AddMatrixRow(int oRow, bool RowIserted)
//        {
//            try
//            {
//                oForm.Freeze(true);
//                if (RowIserted == false)
//                {
//                    oDS_PS_FX010L.InsertRecord((oRow));
//                }

//                oMat01.AddRow();
//                oDS_PS_FX010L.Offset = oRow;
//                oDS_PS_FX010L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

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
//        /// PS_FX010_MTX01
//        /// </summary>
//        private void PS_FX010_MTX01()
//        {
//            string errMessage = string.Empty;

//            short i = 0;
//            string sQry = null;
//            short ErrNum = 0;
//            string SBPLID = null; //사업장
//            string SHisType = null; //이력구분
//            string DocDateF = null; 
//            string DocDateT = null;

//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbouiCOM.ProgressBar ProgressBar01 = null;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                oForm.Freeze(true);
//                SBPLID = oForm.Items.Item("SBPLId").Specific.VALUE.ToString().Trim();
//                SHisType = oForm.Items.Item("SHisType").Specific.VALUE.ToString().Trim();
//                DocDateF = oForm.Items.Item("DocDateF").Specific.VALUE.ToString().Trim();
//                DocDateT = oForm.Items.Item("DocDateT").Specific.VALUE.ToString().Trim();

//                SAPbouiCOM.ProgressBar ProgBar01 = null;
//                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

//                oForm.Freeze(true);

//                sQry = "EXEC [PS_FX010_01] '" + SBPLID + "','" + DocDateF + "','" + DocDateT + "','" + SHisType + "'";
//                oRecordSet01.DoQuery(sQry);

//                oMat01.Clear();
//                oDS_PS_FX010L.Clear();
//                oMat01.FlushToDataSource();
//                oMat01.LoadFromDataSource();

//                if (oRecordSet01.RecordCount == 0)
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                    PS_FX010_AddMatrixRow(0,  true);
//                    PS_FX010_LoadCaption();
//                    errMessage = "조회 결과가 없습니다. 확인하세요.";
//                    throw new Exception();
//                }

//                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
//                {
//                    if (i + 1 > oDS_PS_FX010L.Size)
//                    {
//                        oDS_PS_FX010L.InsertRecord((i));
//                    }

//                    oMat01.AddRow();
//                    oDS_PS_FX010L.Offset = i;

//                    oDS_PS_FX010L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                    oDS_PS_FX010L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("BPLId").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet01.Fields.Item("DocDate").Value).ToString("yyyyMMdd"));
//                    oDS_PS_FX010L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("HIsType").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ClasCode").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("AmtYN").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("FixCode").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("SubCode").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("FixName").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("Qty").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("Amt").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColTxt01", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
//                    oDS_PS_FX010L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());

//                    oRecordSet01.MoveNext();
//                    ProgBar01.Value += 1;
//                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//                }
//                oMat01.LoadFromDataSource();
//                oMat01.AutoResizeColumns();
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
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
//            }
//        }

//        /// <summary>
//        /// DocEntry 초기화
//        /// </summary>
//        private void PS_FX010_FormClear()
//        {
//            string sQry;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]";
//                oRecordSet01.DoQuery(sQry);

//                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) == 0)
//                {
//                    oDS_PS_FX010H.SetValue("DocEntry", 0, "1");
//                }
//                else
//                {
//                    oDS_PS_FX010H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) + 1));
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//        }

//        /// <summary>
//        /// PS_FX010_etBaseForm
//        /// </summary>
//        private void PS_FX010_DeleteData()
//        {
//            string errMessage = string.Empty; 
//            short i = 0;
//            string sQry = null;
//            short ErrNum = 0;
//            string DocEntry = null;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
//            try
//            {
//                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                {
//                    DocEntry = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();

//                    sQry = "SELECT COUNT(*) FROM [@PS_FX010H] WHERE DocEntry = '" + DocEntry + "'";
//                    oRecordSet01.DoQuery(sQry);

//                    if ((oRecordSet01.RecordCount == 0))
//                    {
//                        errMessage = "삭제대상이 없습니다. 확인하세요.";
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        throw new Exception();
//                    }
//                    else
//                    {
//                        sQry = "DELETE FROM [@PS_FX010H] WHERE DocEntry = '" + DocEntry + "'";
//                        oRecordSet01.DoQuery(sQry);
//                    }
//                }
//                dataHelpClass.MDC_GF_Message("삭제 완료!", "S");
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
//        /// PS_FX010_UpdateData
//        /// </summary>
//        private bool PS_FX010_UpdateData()
//        {
//            bool ReturnValue = false;
//            string errMessage = string.Empty;
//            short i = 0;
//            short j = 0;
//            string sQry = null;
//            short DocEntry = 0;
//            string BPLId = null; //사업장
//            string HisType = null; //이력구분
//            string ClasCode = null; //자산구분
//            string DocDate = null; //이력일자
//            string FixCode = null; //자산코드
//            string SubCode = null; //자산순번
//            string FixName = null; //자산명
//            string AmtYN = null; //자본적지출여부
//            string Qty = null; //수량
//            string Amt = null; //금액
//            string Comments = null; //비고사항
//            string CardCode = null; //거래처
//            string CardName = null; //거래처명
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                DocEntry = Convert.ToInt16(oForm.Items.Item("DocEntry").Specific.VALUE);
//                BPLId = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim(); //사업장
//                HisType = oForm.Items.Item("HisType").Specific.VALUE.ToString().Trim(); //이력구분
//                ClasCode = oForm.Items.Item("ClasCode").Specific.VALUE.ToString().Trim(); //자산구분
//                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim(); //이력일자
//                AmtYN = oForm.Items.Item("AmtYN").Specific.VALUE.ToString().Trim(); //자본적지출여부
//                FixCode = oForm.Items.Item("FixCode").Specific.VALUE.ToString().Trim();  //자산코드
//                SubCode = oForm.Items.Item("SubCode").Specific.VALUE.ToString().Trim(); //자산순번
//                FixName = oForm.Items.Item("FixName").Specific.VALUE.ToString().Trim(); //자산명
//                Qty = oForm.Items.Item("Qty").Specific.VALUE.ToString().Trim(); //수량
//                Amt = oForm.Items.Item("Amt").Specific.VALUE.ToString().Trim(); //금액
//                Comments = oForm.Items.Item("Comments").Specific.VALUE.ToString().Trim(); //비고
//                CardCode = oForm.Items.Item("CardCode").Specific.VALUE.ToString().Trim(); //거래처
//                CardName = oForm.Items.Item("CardName").Specific.VALUE.ToString().Trim(); //거래처명

//                if (string.IsNullOrEmpty(Convert.ToString(DocEntry)))
//                {
//                    errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!";
//                    throw new Exception();
//                }

//                sQry = "        UPDATE   [@PS_FX010H]";
//                sQry = sQry + " SET      U_BPLId = '" + BPLId + "',";
//                sQry = sQry + "          U_HisType = '" + HisType + "',";
//                sQry = sQry + "          U_ClasCode = '" + ClasCode + "',";
//                sQry = sQry + "          U_DocDate = '" + DocDate + "',";
//                sQry = sQry + "          U_AmtYN = '" + AmtYN + "',";
//                sQry = sQry + "          U_FixCode = '" + FixCode + "',";
//                sQry = sQry + "          U_SubCode = '" + SubCode + "',";
//                sQry = sQry + "          U_FixName = '" + FixName + "',";
//                sQry = sQry + "          U_Qty  = '" + Qty + "',";
//                sQry = sQry + "          U_Amt  = '" + Amt + "',";
//                sQry = sQry + "          U_Comments = '" + Comments + "',";
//                sQry = sQry + "          U_CardCode = '" + CardCode + "',";
//                sQry = sQry + "          U_CardName = '" + CardName + "'";
//                sQry = sQry + " WHERE    DocEntry = '" + DocEntry + "'";

//                oRecordSet01.DoQuery(sQry);

//                dataHelpClass.MDC_GF_Message("수정 완료!", "S");
//                ReturnValue = true;
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
//            return ReturnValue;
//        }


//        /// <summary>
//        /// PS_FX010_UpdateData
//        /// </summary>
//        private bool PS_FX010_AddData()
//        {
//            bool ReturnValue = false;
//            short i = 0;
//            string sQry = null;
//            short DocEntry = 0;
//            string BPLId = null; //사업장
//            string HisType = null; //이력구분
//            string ClasCode = null; //자산구분
//            string DocDate = null;  //이력일자
//            string FixCode = null; //자산코드
//            string SubCode = null; //자산순번
//            string FixName = null; //자산명
//            string AmtYN = null; //자본적지출여부
//            string Qty = null; //수량
//            string Amt = null; //금액
//            string Comments = null; //비고사항
//            string CardCode = null; //거래처
//            string CardName = null; //거래처명
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                BPLId = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim(); //사업장
//                HisType = oForm.Items.Item("HisType").Specific.VALUE.ToString().Trim(); //이력구분
//                ClasCode = oForm.Items.Item("ClasCode").Specific.VALUE.ToString().Trim(); //자산구분
//                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim(); //이력일자
//                AmtYN = oForm.Items.Item("AmtYN").Specific.VALUE.ToString().Trim(); //자본적지출여부
//                FixCode = oForm.Items.Item("FixCode").Specific.VALUE.ToString().Trim(); //자산코드
//                SubCode = oForm.Items.Item("SubCode").Specific.VALUE.ToString().Trim(); //자산순번
//                FixName = oForm.Items.Item("FixName").Specific.VALUE.ToString().Trim(); //자산명
//                Qty = oForm.Items.Item("Qty").Specific.VALUE.ToString().Trim(); //수량
//                Amt = oForm.Items.Item("Amt").Specific.VALUE.ToString().Trim(); //금액
//                Comments = oForm.Items.Item("Comments").Specific.VALUE.ToString().Trim(); //비고
//                CardCode = oForm.Items.Item("CardCode").Specific.VALUE.ToString().Trim(); //거래처
//                CardName = oForm.Items.Item("CardName").Specific.VALUE.ToString().Trim(); //거래처명

//                //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
//                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_FX010H]";
//                oRecordSet01.DoQuery(sQry);

//                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) == 0)
//                {
//                    DocEntry = 1;
//                }
//                else
//                {
//                    DocEntry = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) + 1;
//                }

//                sQry = "           INSERT INTO [@PS_FX010H]";
//                sQry = sQry + " (";
//                sQry = sQry + "     DocEntry,";
//                sQry = sQry + "     DocNum,";
//                sQry = sQry + "     U_BPLId,";
//                sQry = sQry + "     U_HIsType,";
//                sQry = sQry + "     U_ClasCode,";
//                sQry = sQry + "     U_DocDate,";
//                sQry = sQry + "     U_FixCode,";
//                sQry = sQry + "     U_SubCode,";
//                sQry = sQry + "     U_FixName,";
//                sQry = sQry + "     U_AmtYN,";
//                sQry = sQry + "     U_Qty,";
//                sQry = sQry + "     U_Amt,";
//                sQry = sQry + "     U_Comments,";
//                sQry = sQry + "     U_CardCode,";
//                sQry = sQry + "     U_CardName";
//                sQry = sQry + " )";
//                sQry = sQry + " VALUES";
//                sQry = sQry + " (";
//                sQry = sQry + DocEntry + ",";
//                sQry = sQry + DocEntry + ",";
//                sQry = sQry + "'" + BPLId + "',";
//                sQry = sQry + "'" + HisType + "',";
//                sQry = sQry + "'" + ClasCode + "',";
//                sQry = sQry + "'" + DocDate + "',";
//                sQry = sQry + "'" + FixCode + "',";
//                sQry = sQry + "'" + SubCode + "',";
//                sQry = sQry + "'" + FixName + "',";
//                sQry = sQry + "'" + AmtYN + "',";
//                sQry = sQry + "'" + Qty + "',";
//                sQry = sQry + "'" + Amt + "',";
//                sQry = sQry + "'" + Comments + "',";
//                sQry = sQry + "'" + CardCode + "',";
//                sQry = sQry + "'" + CardName + "'";
//                sQry = sQry + ")";

//                oRecordSet02.DoQuery(sQry);

//                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
//                ReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//            }
//            return ReturnValue;
//        }

//        /// <summary>
//        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
//        /// </summary>
//        /// <param name="oUID"></param>
//        /// <param name="oRow"></param>
//        /// <param name="oCol"></param>
//        private void PS_FX010_FlushToItemValue(string oUID, int oRow, string oCol)
//        {
//            short i = 0;
//            short ErrNum = 0;
//            string sQry = null;
//            string ItemCode = null;
//            short Qty = 0;
//            double Calculate_Weight = 0;
//            string FixCode = null;
//            string SubCode = null;
//            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                switch (oUID)
//                {
//                    case "CFixCode":
//                        FixCode = codeHelpClass.Left(oForm.Items.Item("CFixCode").Specific.VALUE, 6);
//                        SubCode = codeHelpClass.Right(oForm.Items.Item("CFixCode").Specific.VALUE, 3);
//                        oForm.Items.Item("FixCode").Specific.VALUE = FixCode;
//                        oForm.Items.Item("SubCode").Specific.VALUE = SubCode;

//                        sQry = "Select U_FixName From [@PS_FX005H] Where U_FixCode = '" + FixCode + "'";
//                        sQry = sQry + " and U_SubCode = '" + SubCode + "'";
//                        oRecordSet01.DoQuery(sQry);
//                        oForm.Items.Item("FixName").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
//                        break;

//                    case "SubCode": //자산코드
//                        break;

//                    case "CardCode": //거래처
//                        oDS_PS_FX010H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("CardCode").Specific.VALUE + "'"));
//                        break;
//                }

//                if (ErrNum == 1)
//                {
//                    MDC_Com.MDC_GF_Message(ref "시각은 숫자만 입력이 가능합니다.", ref "E");
//                }
//                else if (ErrNum == 2)
//                {
//                    MDC_Com.MDC_GF_Message(ref "시각(시)는 24미만의 값만 입력이 가능합니다.", ref "E");
//                }
//                else if (ErrNum == 3)
//                {
//                    MDC_Com.MDC_GF_Message(ref "시각(분)은 60미만의 값만 입력이 가능합니다.", ref "E");
//                }
//                else
//                {
//                    MDC_Com.MDC_GF_Message(ref "PS_FX010_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//                }


//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//            }
//        }


//        /// <summary>
//        /// FormReset
//        /// </summary>
//        private void PS_FX010_FormReset()
//        {
//            string sQry = string.Empty;
//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
//            try
//            {
//                oForm.Freeze(true);

//                //관리번호
//                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]";
//                oRecordSet.DoQuery(sQry);

//                if (Convert.ToDouble(Strings.Trim(oRecordSet.Fields.Item(0).Value)) == 0)
//                {
//                    oDS_PS_FX010H.SetValue("DocEntry", 0, Convert.ToString(1));
//                }
//                else
//                {
//                    oDS_PS_FX010H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1));
//                }

//                if (string.IsNullOrEmpty(oBPLId))
//                    oBPLId = MDC_PS_Common.User_BPLId();
//                if (string.IsNullOrEmpty(oHisType))
//                    oHisType = "%";
//                if (string.IsNullOrEmpty(oClasCode))
//                    oClasCode = "%";
//                if (string.IsNullOrEmpty(oDocdate))
//                    oDocdate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDocdate, "YYYYMMDD");

//                oDS_PS_FX010H.SetValue("U_BPLId", 0, oBPLId);
//                //사업장
//                oDS_PS_FX010H.SetValue("U_HisType", 0, oHisType);
//                //이력구분
//                oDS_PS_FX010H.SetValue("U_ClasCode", 0, oClasCode);
//                //자산분류
//                oDS_PS_FX010H.SetValue("U_AmtYN", 0, "N");
//                //자본적지출여부
//                oDS_PS_FX010H.SetValue("U_DocDate", 0, oDocdate);
//                //이력일자
//                oDS_PS_FX010H.SetValue("U_FixCode", 0, "");
//                //자산코드
//                oDS_PS_FX010H.SetValue("U_SubCode", 0, "");
//                //자산순번
//                oDS_PS_FX010H.SetValue("U_FixName", 0, "");
//                //자산명
//                oDS_PS_FX010H.SetValue("U_Qty", 0, Convert.ToString(0));
//                //수량
//                oDS_PS_FX010H.SetValue("U_Amt", 0, Convert.ToString(0));
//                //금액
//                oDS_PS_FX010H.SetValue("U_Comments", 0, "");
//                //비고사항

//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("CFixCode").Specific.VALUE = "";

//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//                oForm.Freeze(false);
//            }
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
//                if (pval.BeforeAction == true)
//                {

//                    if (pval.ItemUID == "PS_FX010")
//                    {
//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                        {
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                        {
//                        }
//                    }

//                    ///추가/확인 버튼클릭
//                    if (pval.ItemUID == "BtnAdd")
//                    {

//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {

//                            if (PS_FX010_HeaderSpaceLineDel() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }

//                            if (PS_FX010_AddData() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }

//                            PS_FX010_FormReset();
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                            PS_FX010_LoadCaption();
//                            PS_FX010_MTX01();

//                            oLast_Mode = oForm.Mode;

//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                        {

//                            if (PS_FX010_HeaderSpaceLineDel() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }

//                            if (PS_FX010_UpdateData() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }

//                            PS_FX010_FormReset();
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//                            PS_FX010_LoadCaption();
//                            PS_FX010_MTX01();

//                            //                oForm.Items("GCode").Click ct_Regular
//                        }

//                        ///조회
//                    }
//                    else if (pval.ItemUID == "BtnSearch")
//                    {

//                        PS_FX010_FormReset();
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        ///fm_VIEW_MODE

//                        PS_FX010_LoadCaption();
//                        PS_FX010_MTX01();

//                        ///삭제
//                    }
//                    else if (pval.ItemUID == "BtnDelete")
//                    {

//                        if (SubMain.Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
//                        {

//                            PS_FX010_DeleteData();
//                            PS_FX010_FormReset();
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                            ///fm_VIEW_MODE

//                            PS_FX010_LoadCaption();
//                            PS_FX010_MTX01();

//                        }
//                        else
//                        {

//                        }

//                    }

//                }
//                else if (pval.BeforeAction == false)
//                {
//                    if (pval.ItemUID == "PS_FX010")
//                    {
//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                        {
//                        }
//                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                        {
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
//                if (pval.BeforeAction == true)
//                {
//                    if (pval.CharPressed == 9)
//                    {
//                        if (pval.ItemUID == "CFixCode")
//                        {
//                            //UPGRADE_WARNING: oForm.Items(CFixCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            if (string.IsNullOrEmpty(oForm.Items.Item("CFixCode").Specific.VALUE))
//                            {
//                                SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//                                BubbleEvent = false;
//                            }
//                        }
//                        else if (pval.ItemUID == "CardCode")
//                        {
//                            //UPGRADE_WARNING: oForm.Items(CardCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.VALUE))
//                            {
//                                SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//                                BubbleEvent = false;
//                            }
//                        }
//                    }
//                    //        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "ItemCode", "") '//사용자값활성
//                    //        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
//                }
//                else if (pval.BeforeAction == false)
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
//        /// GOT_FOCUS 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pval.ItemUID == "Mat01")
//                {
//                    if (pval.Row > 0)
//                    {
//                        oLastItemUID01 = pval.ItemUID;
//                        oLastColUID01 = pval.ColUID;
//                        oLastColRow01 = pval.Row;
//                    }
//                }
//                else
//                {
//                    oLastItemUID01 = pval.ItemUID;
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
//                oForm.Freeze(true);
//                if (pval.BeforeAction == true)
//                {

//                }
//                else if (pval.BeforeAction == false)
//                {
//                    if (pval.ItemUID == "BPLId")
//                    {
//                        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oBPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.VALUE);
//                    }

//                    if (pval.ItemUID == "ClasCode")
//                    {
//                        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oClasCode = oForm.Items.Item("ClasCode").Specific.VALUE;
//                    }

//                    if (pval.ItemUID == "ClasCode")
//                    {
//                        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        oHisType = oForm.Items.Item("HisType").Specific.VALUE;
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
//                if (pval.BeforeAction == true)
//                {

//                    if (pval.ItemUID == "Mat01")
//                    {

//                        if (pval.Row > 0)
//                        {

//                            oMat01.SelectRow(pval.Row, true, false);

//                            oForm.Freeze(true);

//                            //DataSource를 이용하여 각 컨트롤에 값을 출력
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pval.Row).Specific.VALUE);
//                            //관리번호
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_BPLId", 0, oMat01.Columns.Item("BPLId").Cells.Item(pval.Row).Specific.VALUE);
//                            //사업장
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_HisType", 0, oMat01.Columns.Item("HisType").Cells.Item(pval.Row).Specific.VALUE);
//                            //이력구분
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_ClasCode", 0, oMat01.Columns.Item("ClasCode").Cells.Item(pval.Row).Specific.VALUE);
//                            //자산분류
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_AmtYN", 0, oMat01.Columns.Item("AmtYN").Cells.Item(pval.Row).Specific.VALUE);
//                            //자본적지출여부
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_DocDate", 0, oMat01.Columns.Item("DocDate").Cells.Item(pval.Row).Specific.VALUE);
//                            //이력일자
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_FixCode", 0, oMat01.Columns.Item("FixCode").Cells.Item(pval.Row).Specific.VALUE);
//                            //자산코드
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_SubCode", 0, oMat01.Columns.Item("SubCode").Cells.Item(pval.Row).Specific.VALUE);
//                            //자산순번
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_FixName", 0, oMat01.Columns.Item("FixName").Cells.Item(pval.Row).Specific.VALUE);
//                            //자산명
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_Qty", 0, oMat01.Columns.Item("Qty").Cells.Item(pval.Row).Specific.VALUE);
//                            //수량
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_Amt", 0, oMat01.Columns.Item("Amt").Cells.Item(pval.Row).Specific.VALUE);
//                            //금액
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_Comments", 0, oMat01.Columns.Item("Comments").Cells.Item(pval.Row).Specific.VALUE);
//                            //비고
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_CardCode", 0, oMat01.Columns.Item("CardCode").Cells.Item(pval.Row).Specific.VALUE);
//                            //거래처
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oDS_PS_FX010H.SetValue("U_CardName", 0, oMat01.Columns.Item("CardName").Cells.Item(pval.Row).Specific.VALUE);
//                            //거래처명

//                            //UPGRADE_WARNING: oForm.Items(CFixCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: oMat01.Columns(SubCode).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("CFixCode").Specific.VALUE = oMat01.Columns.Item("FixCode").Cells.Item(pval.Row).Specific.VALUE + "-" + oMat01.Columns.Item("SubCode").Cells.Item(pval.Row).Specific.VALUE;

//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//                            PS_FX010_LoadCaption();

//                            oForm.Freeze(false);

//                        }
//                    }
//                }
//                else if (pval.BeforeAction == false)
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
//        /// MATRIX_LINK_PRESSED 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
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
//            try
//            {
//                if (pval.BeforeAction == true)
//                {

//                    if (pval.ItemChanged == true)
//                    {

//                        if ((pval.ItemUID == "Mat01"))
//                        {
//                            //                If (pval.ColUID = "ItemCode") Then
//                            //                    '//기타작업
//                            //                    Call oDS_PS_FX010L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
//                            //                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_FX010L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
//                            //                        PS_FX010_AddMatrixRow (pval.Row)
//                            //                    End If
//                            //                Else
//                            //                    Call oDS_PS_FX010L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
//                            //                End If
//                        }
//                        else
//                        {

//                            PS_FX010_FlushToItemValue(pval.ItemUID);

//                            if (pval.ItemUID == "BPLId")
//                            {
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oBPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.VALUE);
//                            }

//                            if (pval.ItemUID == "DocDate")
//                            {
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oDocdate = Strings.Trim(oForm.Items.Item("DocDate").Specific.VALUE);
//                            }

//                            if (pval.ItemUID == "HisType")
//                            {
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oHisType = oForm.Items.Item("HisType").Specific.VALUE;
//                            }

//                            if (pval.ItemUID == "ClasCode")
//                            {
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oClasCode = oForm.Items.Item("ClasCode").Specific.VALUE;
//                            }

//                        }
//                        //            oMat01.LoadFromDataSource
//                        //            oMat01.AutoResizeColumns
//                        //            oForm.Update
//                    }

//                }
//                else if (pval.BeforeAction == false)
//                {

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
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
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
//        /// CHOOSE_FROM_LIST 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
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
//        /// Raise_EVENT_DOUBLE_CLICK 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {

//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
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
//        /// RESIZE 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
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
//        /// EVENT_ROW_DELETE
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                iint i = 0;
//                if ((oLastColRow01 > 0))
//                {
//                    if (pval.BeforeAction == true)
//                    {
//                        //            If (PS_FX010_Validate("행삭제") = False) Then
//                        //                BubbleEvent = False
//                        //                Exit Sub
//                        //            End If
//                        ////행삭제전 행삭제가능여부검사
//                    }
//                    else if (pval.BeforeAction == false)
//                    {
//                        for (i = 1; i <= oMat01.VisualRowCount; i++)
//                        {
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
//                        }
//                        oMat01.FlushToDataSource();
//                        oDS_PS_FX010H.RemoveRecord(oDS_PS_FX010H.Size - 1);
//                        oMat01.LoadFromDataSource();
//                        if (oMat01.RowCount == 0)
//                        {
//                            PS_FX010_AddMatrixRow(0);
//                        }
//                        else
//                        {
//                            if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_FX010H.GetValue("U_CntcCode", oMat01.RowCount - 1))))
//                            {
//                                PS_FX010_AddMatrixRow(oMat01.RowCount);
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
//                            break;
//                        case "1281": //찾기
//                            break;
//                        case "1282": //추가
//                            PS_FX010_FormReset();
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                            BubbleEvent = false;
//                            PS_FX010_LoadCaption();
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
//                            break;
//                        case "1281": //찾기
//                        case "1282": //추가
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
