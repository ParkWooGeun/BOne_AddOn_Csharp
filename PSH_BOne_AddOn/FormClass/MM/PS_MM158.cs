using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 외주자재입고등록
    /// </summary>
    internal class PS_MM158 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_MM158H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM158L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_MM158A;//등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        string LDocEntry; //마지막문서번호
  
        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM158.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM158_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM158");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM158_CreateItems();
                PS_MM158_CF_ChooseFromList();
                PS_MM158_ComboBox_Setting();
                PS_MM158_EnableMenus();
                PS_MM158_SetDocument(oFormDocEntry);
                PS_MM158_MTX01();
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
        private void PS_MM158_CreateItems()
        {
            try
            {
                oDS_PS_MM158A = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oDS_PS_MM158H = oForm.DataSources.DBDataSources.Item("@PS_MM158H");
                oDS_PS_MM158L = oForm.DataSources.DBDataSources.Item("@PS_MM158L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM158_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("ChkYN").Specific.ValidValues.Add("N", "승인대기");
                oForm.Items.Item("ChkYN").Specific.ValidValues.Add("Y", "승인");
                oForm.Items.Item("ChkYN").Specific.ValidValues.Add("C", "승인취소");
                oForm.Items.Item("ChkYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }


        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_MM158_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLWHSCODE2";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oColumn.ChooseFromListUID = "CFLWHSCODE2";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }

                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }

                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }

                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }

                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }

                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }

                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }
            }
        }
        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_MM158_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_MM158_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM158_FormItemEnabled();
                    PS_MM158_AddMatrixRow(0, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM158_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM158_FormClear();
                    oForm.EnableMenu("1281", true);//찾기
                    oForm.EnableMenu("1282", false);//추가
                    oForm.Items.Item("sFocus").Click();
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ChkYN").Enabled = true;
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oMat01.Columns.Item("PP030Doc").Visible = true;
                    oMat01.Columns.Item("ItemCode").Visible = true;
                    oMat01.Columns.Item("ItemSpec").Visible = true;
                    oMat01.Columns.Item("PP030Doc").Editable = true;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemSpec").Editable = true;
                    oMat01.Columns.Item("CADNo").Editable = true;
                    oMat01.Columns.Item("HeatNo").Editable = true;
                    oMat01.Columns.Item("Quantity").Editable = true;
                    oMat01.Columns.Item("Price").Editable = true;
                    oMat01.Columns.Item("Amount").Editable = true;
                    oMat01.Columns.Item("InDate").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                }

                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    if(oForm.Items.Item("ChkYN").Specific.Value.ToString().Trim() == "Y" || oForm.Items.Item("ChkYN").Specific.Value.ToString().Trim() == "C")
                    {
                        oForm.EnableMenu("1281", true); //찾기
                        oForm.EnableMenu("1282", true); //추가
                        oForm.Items.Item("sFocus").Click();
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("ChkYN").Enabled = true; 
                        oMat01.Columns.Item("PP030Doc").Editable = false;
                        oMat01.Columns.Item("ItemCode").Editable = false;
                        oMat01.Columns.Item("ItemSpec").Editable = false;
                        oMat01.Columns.Item("CADNo").Editable = false;
                        oMat01.Columns.Item("HeatNo").Editable = false;
                        oMat01.Columns.Item("Quantity").Editable = false;
                        oMat01.Columns.Item("Price").Editable = false;
                        oMat01.Columns.Item("Amount").Editable = false;
                        oMat01.Columns.Item("InDate").Editable = false;
                    }
                    
                    else
                    {
                        oForm.EnableMenu("1281", true); //찾기
                        oForm.EnableMenu("1282", true); //추가
                        oForm.Items.Item("sFocus").Click();
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("ChkYN").Enabled = true;
                        oMat01.Columns.Item("PP030Doc").Editable = false;
                        oMat01.Columns.Item("ItemCode").Editable = false;
                        oMat01.Columns.Item("ItemSpec").Editable = true;
                        oMat01.Columns.Item("CADNo").Editable = true;
                        oMat01.Columns.Item("HeatNo").Editable = true;
                        oMat01.Columns.Item("Quantity").Editable = true;
                        oMat01.Columns.Item("Price").Editable = true;
                        oMat01.Columns.Item("Amount").Editable = true;
                        oMat01.Columns.Item("InDate").Editable = true;
                    }
                }
                oMat01.AutoResizeColumns();
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
        /// PS_MM158_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM158_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM158L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM158L.Offset = oRow;
                oDS_PS_MM158L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// PS_MM158_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM158A_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM158A.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_MM158A.Offset = oRow;
                oDS_PS_MM158A.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM158_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM158'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
        /// <summary>
		/// PS_MM158_Grid01
		/// </summary>
		private void PS_MM158_MTX01()
        {
            string sQry = string.Empty;
            string errMessage = string.Empty;
            int loopCount;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try

            {
                oForm.Freeze(true);

                sQry = "SELECT MM158H.DocEntry AS 문서번호, MM158H.U_CardCode AS 거래처코드, MM158H.U_CardName AS 거래처명, ";
                sQry += " Convert(Char(10),MM158H.U_DocDate,23) AS 전기일, MM158L.U_PP030Doc AS 작번, MM158L.U_ItemCode AS 품목번호,";
                sQry += " MM158L.U_ItemName AS 품목이름, MM158L.U_ItemSpec AS 규격, ISNULL(MM158L.U_CADNo,'') AS 도면No, ";
                sQry += " ISNULL(MM158L.U_HeatNo,'') AS HeatNo, MM158L.U_Quantity AS 입고량, MM158L.U_FQty AS 예상수량,";
                sQry += "MM158L.U_Price AS 단가, MM158L.U_Amount AS 총금액";
                sQry += " FROM [@PS_MM158H] MM158H INNER JOIN [@PS_MM158L] MM158L ON MM158H.DocEntry = MM158L.DocEntry";
                sQry += " WHERE MM158H.Canceled = 'N' AND MM158H.U_ChkYN = 'N'";
                oRecordSet.DoQuery(sQry);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                 for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_MM158A.InsertRecord(loopCount);
                    }
                    oDS_PS_MM158A.Offset = loopCount;

                    oDS_PS_MM158A.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                              
                    oDS_PS_MM158A.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("문서번호").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("거래처코드").Value.ToString().Trim());
                    oDS_PS_MM158A.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("거래처명").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("전기일").Value.ToString().Trim());    
                    oDS_PS_MM158A.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("작번").Value.ToString().Trim());    
                    oDS_PS_MM158A.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("품목번호").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("품목이름").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("규격").Value.ToString().Trim());
                    oDS_PS_MM158A.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("도면No").Value.ToString().Trim());   
                    oDS_PS_MM158A.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("HeatNo").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("입고량").Value.ToString().Trim());  
                    oDS_PS_MM158A.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("예상수량").Value.ToString().Trim());
                    oDS_PS_MM158A.SetValue("U_ColReg13", loopCount, oRecordSet.Fields.Item("단가").Value.ToString().Trim());
                    oDS_PS_MM158A.SetValue("U_ColReg14", loopCount, oRecordSet.Fields.Item("총금액").Value.ToString().Trim());

                    oRecordSet.MoveNext();
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
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
                oForm.Freeze(false);
            }
        }
        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_MM158_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oMat01.FlushToDataSource();
                // Matrix 마지막 행 삭제(DB 저장시)
                //if (oDS_PS_MM158L.Size > 1)
                if(string.IsNullOrEmpty(oMat01.Columns.Item("PP030Doc").Cells.Item(oDS_PS_MM158L.Size).Specific.Value))
                {
                    oDS_PS_MM158L.RemoveRecord(oDS_PS_MM158L.Size - 1);
                }
                oMat01.LoadFromDataSource();

                if (string.IsNullOrEmpty(oForm.Items.Item("CardName").Specific.Value))
                {
                    errMessage = "거래처가 입력되지 않았습니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "전기일 입력되지 않았습니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("ChkYN").Specific.Value))
                {
                    errMessage = "승인구분 입력되지 않았습니다.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount < 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (int i = 1; i <= oMat01.VisualRowCount ; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030Doc").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작번은 필수입니다.";
                        oMat01.Columns.Item("PP030Doc").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목번호은 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목내역은 필수입니다.";
                        oMat01.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    else if(Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value) < 0)
                    {
                        errMessage = "입고량은 필수입니다.";
                        oMat01.Columns.Item("Quantity").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if(string.IsNullOrEmpty(oMat01.Columns.Item("InDate").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "입고일은 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if(string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고선택은 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if(string.IsNullOrEmpty(oMat01.Columns.Item("WhsName").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고선택은 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oMat01.Columns.Item("OrQty").Cells.Item(i).Specific.Value) < Convert.ToDouble(oMat01.Columns.Item("FQty").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "예상수량은 수주수량을 초과할 수 없습니다..";
                        oMat01.Columns.Item("FQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }


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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// 입고DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM158_DI_API()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            int LineNumCount;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                LineNumCount = 0;
                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }
                oDIObject.UserFields.Fields.Item("Comments").Value = "애드온 입고 문서번호:" + oForm.Items.Item("DocEntry").Specific.Value + " 외주자재입고_PS_MM158";
                
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(LineNumCount);
                    oDIObject.Lines.ItemCode = oDS_PS_MM158L.GetValue("U_ItemCode", i).ToString().Trim();
                    oDIObject.Lines.WarehouseCode = oDS_PS_MM158L.GetValue("U_WhsCode", i).ToString().Trim();
                    oDIObject.Lines.Quantity = Convert.ToDouble(oDS_PS_MM158L.GetValue("U_Quantity", i).ToString().Trim());
                    oDIObject.Lines.Price = Convert.ToDouble(0);
                    oDIObject.Lines.LineTotal = Convert.ToDouble(0);
                    oDIObject.Lines.UserFields.Fields.Item("PriceBefDi").Value = Convert.ToDouble(0);
                    oDIObject.Lines.UserFields.Fields.Item("U_OrdNum").Value = oDS_PS_MM158L.GetValue("U_PP030Doc", i).ToString().Trim();
                    oDIObject.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM158L.GetValue("U_HeatNo", i).ToString().Trim();
                    LineNumCount += 1;
                }
                
                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        oMat01.Columns.Item("InDoc").Cells.Item(i).Specific.Value = Convert.ToString(afterDIDocNum);
                        oMat01.Columns.Item("InNum").Cells.Item(i).Specific.Value = i;
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else if (errCode == "3")
                {
                    //PS_MM180_InterfaceB1toR3에서 오류 발생하면 해당 메소드에서 오류 메시지 출력, 이 분기문에서는 별도 메시지 출력 안함
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 출고DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM158_DI_API2()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            int LineNumCount;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                LineNumCount = 0;
                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }
                oDIObject.UserFields.Fields.Item("Comments").Value = "애드온 입고 문서번호:" + oForm.Items.Item("DocEntry").Specific.Value + " 외주자재입고_PS_MM158 (입고취소)";

                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(LineNumCount);
                    oDIObject.Lines.ItemCode = oDS_PS_MM158L.GetValue("U_ItemCode", i).ToString().Trim();
                    oDIObject.Lines.WarehouseCode = oDS_PS_MM158L.GetValue("U_WhsCode", i).ToString().Trim();
                    oDIObject.Lines.Quantity = Convert.ToDouble(oDS_PS_MM158L.GetValue("U_Quantity", i).ToString().Trim());
                    oDIObject.Lines.Price = Convert.ToDouble(0);
                    oDIObject.Lines.LineTotal = Convert.ToDouble(0);
                    oDIObject.Lines.UserFields.Fields.Item("PriceBefDi").Value = Convert.ToDouble(0);
                    oDIObject.Lines.UserFields.Fields.Item("U_OrdNum").Value = oDS_PS_MM158L.GetValue("U_PP030Doc", i).ToString().Trim();
                    oDIObject.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM158L.GetValue("U_HeatNo", i).ToString().Trim();
                    LineNumCount += 1;
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        oMat01.Columns.Item("InDoc").Cells.Item(i).Specific.Value = "";
                        oMat01.Columns.Item("InNum").Cells.Item(i).Specific.Value = "";
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else if (errCode == "3")
                {
                    //PS_MM180_InterfaceB1toR3에서 오류 발생하면 해당 메소드에서 오류 메시지 출력, 이 분기문에서는 별도 메시지 출력 안함
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
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

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (oForm.Items.Item("ChkYN").Specific.Value.ToString().Trim() == "Y")
                            {
                                errMessage = "신규문서는 승인으로 등록할 수 없습니다.";
                                BubbleEvent = false;
                                throw new System.Exception();
                            }

                            if (PS_MM158_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            LDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            //승인된 건은 취소불가
                            if (oForm.Items.Item("ChkYN").Specific.Value.ToString().Trim() == "C")
                            {
                                sQry = " select COUNT(*) FROM [@PS_MM130H] MM130H INNER JOIN [@PS_MM130L] MM130L ON MM130H.DocEntry= MM130L.DocEntry ";
                                sQry += "INNER JOIN(SELECT B.U_PP030Doc, b.U_ItemCode, B.U_HeatNo FROM [@PS_MM158H] A INNER JOIN[@PS_MM158L] B ON A.DocEntry = B.DocEntry WHERE A.Canceled = 'N' AND A.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "')g";
                                sQry += " ON MM130L.U_OrdNum = g.U_PP030Doc AND MM130L.U_OutItmCd = g.U_ItemCode  AND MM130L.U_HeatNo = g.U_HeatNo";
                                sQry += " WHERE MM130H.Canceled = 'N' AND MM130H.U_OKYNC = 'Y'";
                                oRecordSet.DoQuery(sQry);
                                if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
                                { 
                                    errMessage = "반출이 일어난 경우 취소가 불가능합니다.";
                                    BubbleEvent = false;
                                    throw new System.Exception();
                                }
                                else
                                {
                                    if (PS_MM158_DI_API2() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }

                                }

                                //if (string.IsNullOrEmpty(oMat01.Columns.Item("InDoc").Cells.Item(1).Specific.Value))
                                //{
                                //    sQry = "UPDATE [@PS_MM158H] SET Canceled = 'Y' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                //    oRecordSet.DoQuery(sQry);
                                //}
                                //else
                                //{
                                //    errMessage = "반출이 일어난 경우 취소가 불가능합니다.";
                                //    BubbleEvent = false;
                                //    throw new System.Exception();
                                //}
                            }
                            else if (oForm.Items.Item("ChkYN").Specific.Value.ToString().Trim() == "Y")
                            {
                                if (PS_MM158_DataValidCheck() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                sQry = "select count(*) from [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code WHERE A.Code ='MM158' AND B.U_UseYN ='Y' AND B.U_AppUser ='" + PSH_Globals.oCompany.UserName + "'";
                                oRecordSet.DoQuery(sQry);
                                //시스템 코드등록에 자재담당자로 등록되어있고, 입고문서가 없으면 승인처리 가능
                                if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1" && string.IsNullOrEmpty(oMat01.Columns.Item("InDoc").Cells.Item(1).Specific.Value))
                                {
                                    if (PS_MM158_DI_API() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                else
                                {
                                    errMessage = "자재담당자만 승인으로 등록할 수 있습니다.";
                                    BubbleEvent = false;
                                    throw new System.Exception();
                                }
                            }
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
                                PS_MM158_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM158_FormItemEnabled();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_MM158_MTX01();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
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
            oForm.Freeze(true);
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "InDate")
                            {
                                oDS_PS_MM158L.SetValue("U_InDate", pVal.Row - 1, DateTime.Now.ToString("yyyyMMdd"));
                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP030Doc");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ItemCode");
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
        /// Raise_EVENT_DOUBLE_CLICK
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    string DocCode;
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat02.FlushToDataSource();
                        }
                        else
                        {
                            DocCode = oMat02.Columns.Item("DocCode").Cells.Item(pVal.Row).Specific.Value;
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_MM158_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = DocCode;
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
            finally
            {
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    double sum = 0;
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.FlushToDataSource();
                            if (pVal.ColUID == "PP030Doc")
                            {
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM158L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM158_AddMatrixRow(pVal.Row, false);
                                }
                                oDS_PS_MM158L.SetValue("U_OrQty", pVal.Row - 1, dataHelpClass.GetValue("SELECT RDR1.Quantity FROM ORDR ORDR INNER JOIN RDR1 RDR1 ON ORDR.DocEntry = RDR1.DocEntry WHERE ORDR.CANCELED <> 'Y' AND ORDR.DocStatus <> 'C' AND RDR1.ItemCode = '" + oDS_PS_MM158L.GetValue("U_PP030Doc", pVal.Row - 1).ToString().Trim() + "'", 0, 1));
                                oMat01.LoadFromDataSource();
                            }
                            else if (pVal.ColUID == "Quantity")
                            {
                                sum = Double.Parse(oDS_PS_MM158L.GetValue("U_Price", pVal.Row - 1)) * Double.Parse(oDS_PS_MM158L.GetValue("U_Quantity", pVal.Row - 1));
                                oDS_PS_MM158L.SetValue("U_Amount", pVal.Row - 1, sum.ToString());
                            }
                            else if (pVal.ColUID == "Price")
                            {
                                sum = Double.Parse(oDS_PS_MM158L.GetValue("U_Price", pVal.Row - 1)) * Double.Parse(oDS_PS_MM158L.GetValue("U_Quantity", pVal.Row - 1));
                                oDS_PS_MM158L.SetValue("U_Amount",pVal.Row - 1, sum.ToString());
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                oDS_PS_MM158L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.GetValue("SELECT ItemName FROM OITM WHERE ItemCode = '" + oDS_PS_MM158L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim() + "'", 0, 1));
                                oDS_PS_MM158L.SetValue("U_ItemSpec", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_Size FROM OITM WHERE ItemCode = '" + oDS_PS_MM158L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim() + "'", 0, 1));
                                oDS_PS_MM158L.SetValue("U_WhsCode", pVal.Row - 1, "902");
                                oDS_PS_MM158L.SetValue("U_WhsName", pVal.Row - 1, dataHelpClass.GetValue("SELECT WhsName FROM OWHS where WhsCode = '902'", 0, 1));
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "CardCode")
                        {
                            oDS_PS_MM158H.SetValue("U_CardName", 0, dataHelpClass.GetValue("SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'", 0, 1));
                        }
                    }
                    
                }
                else if (pVal.Before_Action == false)
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
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
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
                    //PS_MM158_FormItemEnabled();
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM158H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM158L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                            PS_MM158_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PS_MM158_MTX01();
                            break;
                        case "1282": //추가
                            PS_MM158_FormClear();
                            PS_MM158_FormItemEnabled();
                            PS_MM158_AddMatrixRow(0, true);
                            PS_MM158_MTX01();
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_MM158_FormItemEnabled();
                            PS_MM158_MTX01();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
            finally
            {
            }
        }
    }
}
